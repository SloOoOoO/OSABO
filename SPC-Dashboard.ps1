#Requires -Version 5.1
<#
.SYNOPSIS
    SPC Server Dashboard - zero-dependency WPF monitor for SPC servers.
.DESCRIPTION
    Reads the server list directly from the .xlsm (ZIP+XML, no Excel needed),
    checks servers in parallel, and launches UltraVNC connections.
    Edit the $Config block below to customise paths and settings.
#>

# ============================ CONFIGURATION ==================================
$script:Config = @{
    # Path to the Excel workbook (defaults to the first .xlsm next to this script)
    ExcelFile          = Join-Path $PSScriptRoot "SPC-Geraeteliste_Gesamt_NEU.xlsm"
    # Sheet index (1-based). Sheet 1 = "Handmessplaetze 09.08.2023"
    SheetIndex         = 1
    # First data row (rows 1-2 are headers)
    DataStartRow       = 3
    # UltraVNC executable path
    VncExe             = "C:\Legacy\UltraVNC\UltraVNC-Viewer.exe"
    # VNC password - DO NOT commit a real password; set it here or via env var $env:VNC_PASSWORD
    VncPassword        = if ($env:VNC_PASSWORD) { $env:VNC_PASSWORD } else { "CHANGE_ME" }
    # Auto-refresh interval in seconds (0 = disabled)
    AutoRefreshSeconds = 30
    # Per-host ping timeout in milliseconds
    PingTimeoutMs      = 1500
    # Fallback VNC port check when ICMP is blocked
    TcpFallbackPort    = 5900
    # Fallback TCP timeout in milliseconds
    TcpFallbackTimeoutMs = 500
    # Excel column containing MAC addresses for Wake-on-LAN (default: AA / column 27)
    MacAddressColumn   = 'AA'
}
# ============================================================================

Set-StrictMode -Off
Add-Type -AssemblyName PresentationFramework, PresentationCore, WindowsBase
Add-Type -AssemblyName System.IO.Compression.FileSystem
Add-Type -AssemblyName System.Net.NetworkInformation -ErrorAction SilentlyContinue

# --- ServerItem: INotifyPropertyChanged so WPF bindings update automatically ---
if (-not ('ServerItem' -as [type])) {
    Add-Type @"
using System;
using System.Collections;
using System.ComponentModel;

public class ServerItem : INotifyPropertyChanged {
    private string _status  = "checking";
    private long   _latency = -1L;

    public string    DisplayName { get; set; }
    public string    Hostname    { get; set; }
    public int       RowNumber   { get; set; }
    public Hashtable RawData     { get; set; }

    public string Status {
        get { return _status; }
        set {
            _status = value;
            Fire("Status"); Fire("StatusColor"); Fire("LatencyText");
        }
    }

    public long LatencyMs {
        get { return _latency; }
        set {
            _latency = value;
            Fire("LatencyMs"); Fire("LatencyText");
        }
    }

    // Returns hex colour string used by DataTrigger in XAML
    public string StatusColor {
        get {
            if (_status == "online")  return "#34C759";
            if (_status == "offline") return "#FF3B30";
            return "#8E8E93";
        }
    }

    public string LatencyText {
        get {
            if (_status == "checking") return "-";
            if (_status == "offline" || _latency < 0) return "offline";
            return _latency.ToString() + " ms";
        }
    }

    public event PropertyChangedEventHandler PropertyChanged;
    private void Fire(string n) {
        if (PropertyChanged != null)
            PropertyChanged(this, new PropertyChangedEventArgs(n));
    }
}
"@
}

# ====================== EXCEL PARSING (ZIP + XML) ============================
function Resolve-ExcelFilePath {
    $configuredPath = $script:Config.ExcelFile
    if ($configuredPath -and (Test-Path -LiteralPath $configuredPath)) {
        return $configuredPath
    }

    $fallback = Get-ChildItem -LiteralPath $PSScriptRoot -Filter '*.xlsm' -File -ErrorAction SilentlyContinue |
        Select-Object -First 1
    if ($fallback) {
        return $fallback.FullName
    }

    return $configuredPath
}

function ConvertFrom-ColLetter ([string]$letters) {
    $letters = $letters.ToUpper()
    $idx = 0
    foreach ($ch in $letters.ToCharArray()) {
        $idx = $idx * 26 + ([int][char]$ch - [int][char]'A' + 1)
    }
    return $idx - 1   # 0-based
}

function Get-ConfiguredColumnIndex([object]$Value, [int]$DefaultIndex) {
    if ($null -eq $Value) { return $DefaultIndex }

    $text = [string]$Value
    if ([string]::IsNullOrWhiteSpace($text)) { return $DefaultIndex }

    $text = $text.Trim()
    if ($text -match '^\d+$') {
        return [int]$text
    }
    if ($text -match '^[A-Za-z]+$') {
        return ConvertFrom-ColLetter $text
    }

    Write-Warning "Invalid column mapping '$text'. Falling back to column index $DefaultIndex."
    return $DefaultIndex
}

$script:MacAddressColumnIndex = Get-ConfiguredColumnIndex $script:Config.MacAddressColumn 26

function Read-ExcelServers {
    param([string]$Path)

    if (-not (Test-Path $Path)) {
        Write-Warning "Excel file not found: $Path"
        return @()
    }

    # Copy to temp so we can read even when Excel has it open
    $tmp = [IO.Path]::GetTempFileName() + '.xlsm'
    try {
        Copy-Item -LiteralPath $Path -Destination $tmp -Force -ErrorAction Stop
    } catch {
        Write-Warning "Cannot copy Excel file (locked?): $_"
        return @()
    }

    try {
        $zip = [IO.Compression.ZipFile]::OpenRead($tmp)

        # -- Shared strings --------------------------------------------------
        $ss = @()
        $ssEntry = $zip.Entries | Where-Object FullName -eq 'xl/sharedStrings.xml' | Select-Object -First 1
        if ($ssEntry) {
            $rdr = New-Object IO.StreamReader($ssEntry.Open())
            [xml]$ssXml = $rdr.ReadToEnd(); $rdr.Close()
            $ns = New-Object Xml.XmlNamespaceManager($ssXml.NameTable)
            $ns.AddNamespace('s','http://schemas.openxmlformats.org/spreadsheetml/2006/main')
            foreach ($si in $ssXml.SelectNodes('//s:si',$ns)) {
                $ss += ($si.SelectNodes('.//s:t',$ns) | ForEach-Object { $_.InnerText }) -join ''
            }
        }

        # -- Sheet XML -------------------------------------------------------
        $sheetName = "xl/worksheets/sheet$($script:Config.SheetIndex).xml"
        $shEntry   = $zip.Entries | Where-Object FullName -eq $sheetName | Select-Object -First 1
        if (-not $shEntry) { $zip.Dispose(); return @() }

        $rdr = New-Object IO.StreamReader($shEntry.Open())
        [xml]$shXml = $rdr.ReadToEnd(); $rdr.Close()
        $zip.Dispose()

        $ns = New-Object Xml.XmlNamespaceManager($shXml.NameTable)
        $ns.AddNamespace('s','http://schemas.openxmlformats.org/spreadsheetml/2006/main')

        # Helper: get typed cell value
        $getCellVal = {
            param($cell)
            $t     = $cell.GetAttribute('t')
            $vNode = $cell.SelectSingleNode('s:v', $ns)
            if (-not $vNode) { return '' }
            $v = $vNode.InnerText
            if ($t -eq 's') {
                $i = [int]$v
                if ($i -lt $ss.Count) { return $ss[$i] } else { return $v }
            }
            return $v
        }

        $servers = [System.Collections.Generic.List[object]]::new()

        foreach ($row in $shXml.SelectNodes('//s:sheetData/s:row', $ns)) {
            $r = [int]$row.GetAttribute('r')
            if ($r -lt $script:Config.DataStartRow) { continue }

            $rowData = @{}
            foreach ($cell in $row.SelectNodes('s:c', $ns)) {
                $ref     = $cell.GetAttribute('r')
                $colLet  = [regex]::Match($ref, '[A-Za-z]+').Value
                $colIdx  = ConvertFrom-ColLetter $colLet
                $rowData[$colIdx] = & $getCellVal $cell
            }

            $colA = if ($rowData.ContainsKey(0))  { $rowData[0]  } else { '' }
            $colZ = if ($rowData.ContainsKey(25)) { $rowData[25] } else { '' }

            # Skip blank rows
            if (-not $colA -and -not $colZ) { continue }

            $item = New-Object ServerItem
            $item.DisplayName = $colA
            $item.Hostname    = $colZ
            $item.RowNumber   = $r
            $item.RawData     = $rowData
            $servers.Add($item) | Out-Null
        }

        return $servers
    } catch {
        Write-Warning "Error reading Excel: $_"
        return @()
    } finally {
        Remove-Item -LiteralPath $tmp -ErrorAction SilentlyContinue
    }
}

# ========================== PARALLEL PING ===================================
function Start-PingAll {
    param([System.Collections.ObjectModel.ObservableCollection[object]]$Collection)

    # Mark all as checking
    foreach ($item in $Collection) {
        $item.Status = 'checking'
        $item.LatencyMs = -1L
    }

    if ($script:PingRunspace) {
        try { $script:PingRunspace.Close(); $script:PingRunspace.Dispose() } catch {}
    }
    if ($script:PingPS) {
        try { $script:PingPS.Dispose() } catch {}
    }

    $script:PingQueue = [System.Collections.Concurrent.ConcurrentQueue[hashtable]]::new()
    Update-StatusBar

    # Build a plain array to pass to the runspace (avoid serialization issues)
    $hostList = @($Collection | Where-Object { $_.Hostname } |
                    ForEach-Object { @{ Host = $_.Hostname } })

    $rs = [System.Management.Automation.Runspaces.RunspaceFactory]::CreateRunspace()
    $rs.Open()
    $rs.SessionStateProxy.SetVariable('hostList',   $hostList)
    $rs.SessionStateProxy.SetVariable('timeoutMs',  $script:Config.PingTimeoutMs)
    $rs.SessionStateProxy.SetVariable('tcpPort',    $script:Config.TcpFallbackPort)
    $rs.SessionStateProxy.SetVariable('tcpTimeoutMs', $script:Config.TcpFallbackTimeoutMs)
    $rs.SessionStateProxy.SetVariable('resultQueue',$script:PingQueue)

    $ps = [System.Management.Automation.PowerShell]::Create()
    $ps.Runspace = $rs
    [void]$ps.AddScript({
        Add-Type -AssemblyName System.Net.NetworkInformation -ErrorAction SilentlyContinue
        $tasks = [System.Collections.Generic.List[object]]::new()
        foreach ($entry in $hostList) {
            $ping = [System.Net.NetworkInformation.Ping]::new()
            try {
                $tasks.Add([pscustomobject]@{
                    Task = $ping.SendPingAsync($entry.Host, $timeoutMs)
                    Ping = $ping
                    Host = $entry.Host
                })
            } catch {
                $resultQueue.Enqueue(@{
                    Host    = $entry.Host
                    Status  = 'offline'
                    Latency = -1L
                })
                $ping.Dispose()
            }
        }
        $tcpTasks = [System.Collections.Generic.List[object]]::new()
        foreach ($t in $tasks) {
            try {
                $reply = $t.Task.GetAwaiter().GetResult()
            } catch {
                $reply = $null
            }
            if ($reply -and $reply.Status -eq [System.Net.NetworkInformation.IPStatus]::Success) {
                $resultQueue.Enqueue(@{
                    Host    = $t.Host
                    Status  = 'online'
                    Latency = [int64]$reply.RoundtripTime
                })
            } else {
                $client = New-Object System.Net.Sockets.TcpClient
                $stopwatch = [System.Diagnostics.Stopwatch]::StartNew()
                $tcpTasks.Add([pscustomobject]@{
                    Host      = $t.Host
                    Client    = $client
                    Connect   = $client.ConnectAsync($t.Host, $tcpPort)
                    Stopwatch = $stopwatch
                })
            }
            $t.Ping.Dispose()
        }
        foreach ($t in $tcpTasks) {
            $isOnline = $false
            $latency = -1L
            try {
                if ($t.Connect.Wait($tcpTimeoutMs) -and $t.Client.Connected) {
                    $isOnline = $true
                    $latency = [int64][Math]::Max(1, [Math]::Round($t.Stopwatch.Elapsed.TotalMilliseconds))
                }
            } catch {}
            $t.Stopwatch.Stop()
            try { $t.Client.Close() } catch {}
            try { $t.Client.Dispose() } catch {}
            $resultQueue.Enqueue(@{
                Host    = $t.Host
                Status  = if ($isOnline) { 'online' } else { 'offline' }
                Latency = $latency
            })
        }
    })

    $script:PingPS = $ps
    $script:PingRunspace = $rs
    $script:PingAsyncResult = $ps.BeginInvoke()
}

# =============================== COLUMN MAP ==================================
$script:ColumnDefs = @(
    @{ Label = 'SPC-Nr.';              Index = 0  }
    @{ Label = 'Typ';                  Index = 1  }
    @{ Label = 'Teil-Name / Nr.';      Index = 2  }
    @{ Label = 'Op.';                  Index = 3  }
    @{ Label = 'Modell';               Index = 4  }
    @{ Label = 'Halle';                Index = 5  }
    @{ Label = 'Pfeiler';              Index = 6  }
    @{ Label = 'Tisch-Nr.';            Index = 7  }
    @{ Label = 'PTM';                  Index = 8  }
    @{ Label = 'Serien-Nr.';           Index = 9  }
    @{ Label = 'Pruefplan / Messprog.'; Index = 10 }
    @{ Label = 'Kanal-Belegung Anz.';  Index = 11 }
    @{ Label = 'Kanal belegt';         Index = 12 }
    @{ Label = 'Sonder-HW';            Index = 13 }
    @{ Label = 'Anwahlbox';            Index = 14 }
    @{ Label = 'Windows-Version';      Index = 15 }
    @{ Label = 'IR-57 Sichtpruefung';  Index = 16 }
    @{ Label = 'Filter / Luefter';     Index = 17 }
    @{ Label = 'NW-faehig';            Index = 18 }
    @{ Label = 'Anzahl Pruefplan';     Index = 19 }
    @{ Label = 'Daten gesichert';      Index = 20 }
    @{ Label = 'Daten upload';         Index = 21 }
    @{ Label = 'Datenmanager';         Index = 22 }
    @{ Label = 'Greengate';            Index = 23 }
    @{ Label = 'Bemerkung';            Index = 24 }
    @{ Label = 'Hostname';             Index = 25 }
    @{ Label = 'MAC-Adresse';          Index = $script:MacAddressColumnIndex }
    @{ Label = 'Serien-Nr. USB';       Index = 27 }
    @{ Label = 'Green Gate korrekt';   Index = 28 }
)

# ================================== XAML =====================================
[xml]$xaml = @'
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="SPC Server Dashboard"
        Width="1120" Height="740"
        MinWidth="820" MinHeight="520"
        Background="#F2F2F7"
        FontFamily="Segoe UI"
        WindowStartupLocation="CenterScreen"
        UseLayoutRounding="True">

  <Window.Resources>

    <!-- Primary (blue) button -->
    <Style x:Key="PrimaryBtn" TargetType="Button">
      <Setter Property="Foreground"       Value="White"/>
      <Setter Property="Background"       Value="#007AFF"/>
      <Setter Property="FontSize"         Value="13"/>
      <Setter Property="Padding"          Value="18,7"/>
      <Setter Property="BorderThickness"  Value="0"/>
      <Setter Property="Cursor"           Value="Hand"/>
      <Setter Property="Template">
        <Setter.Value>
          <ControlTemplate TargetType="Button">
            <Border x:Name="Bd" CornerRadius="8"
                    Background="{TemplateBinding Background}"
                    Padding="{TemplateBinding Padding}">
              <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>
            </Border>
            <ControlTemplate.Triggers>
              <Trigger Property="IsMouseOver" Value="True">
                <Setter TargetName="Bd" Property="Background" Value="#0062CC"/>
              </Trigger>
              <Trigger Property="IsPressed" Value="True">
                <Setter TargetName="Bd" Property="Background" Value="#004EA3"/>
              </Trigger>
              <Trigger Property="IsEnabled" Value="False">
                <Setter TargetName="Bd" Property="Background" Value="#C7C7CC"/>
              </Trigger>
            </ControlTemplate.Triggers>
          </ControlTemplate>
        </Setter.Value>
      </Setter>
    </Style>

    <!-- Secondary (gray) button -->
    <Style x:Key="SecondaryBtn" TargetType="Button">
      <Setter Property="Foreground"       Value="#1C1C1E"/>
      <Setter Property="Background"       Value="#E5E5EA"/>
      <Setter Property="FontSize"         Value="13"/>
      <Setter Property="Padding"          Value="14,7"/>
      <Setter Property="BorderThickness"  Value="0"/>
      <Setter Property="Cursor"           Value="Hand"/>
      <Setter Property="Template">
        <Setter.Value>
          <ControlTemplate TargetType="Button">
            <Border x:Name="Bd" CornerRadius="8"
                    Background="{TemplateBinding Background}"
                    Padding="{TemplateBinding Padding}">
              <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>
            </Border>
            <ControlTemplate.Triggers>
              <Trigger Property="IsMouseOver" Value="True">
                <Setter TargetName="Bd" Property="Background" Value="#D1D1D6"/>
              </Trigger>
              <Trigger Property="IsPressed" Value="True">
                <Setter TargetName="Bd" Property="Background" Value="#AEAEB2"/>
              </Trigger>
            </ControlTemplate.Triggers>
          </ControlTemplate>
        </Setter.Value>
      </Setter>
    </Style>

    <!-- Search TextBox -->
    <Style x:Key="SearchStyle" TargetType="TextBox">
      <Setter Property="FontSize"         Value="13"/>
      <Setter Property="Foreground"       Value="#1C1C1E"/>
      <Setter Property="Background"       Value="#FFFFFF"/>
      <Setter Property="BorderBrush"      Value="#D1D1D6"/>
      <Setter Property="BorderThickness"  Value="1"/>
      <Setter Property="Padding"          Value="10,7"/>
      <Setter Property="Template">
        <Setter.Value>
          <ControlTemplate TargetType="TextBox">
            <Border CornerRadius="8"
                    Background="{TemplateBinding Background}"
                    BorderBrush="{TemplateBinding BorderBrush}"
                    BorderThickness="{TemplateBinding BorderThickness}">
              <ScrollViewer x:Name="PART_ContentHost" Margin="{TemplateBinding Padding}"/>
            </Border>
          </ControlTemplate>
        </Setter.Value>
      </Setter>
    </Style>

    <!-- ListView item container -->
    <Style x:Key="ListItem" TargetType="ListViewItem">
      <Setter Property="HorizontalContentAlignment" Value="Stretch"/>
      <Setter Property="Padding"    Value="0"/>
      <Setter Property="Margin"     Value="0"/>
      <Setter Property="Background" Value="Transparent"/>
      <Setter Property="Template">
        <Setter.Value>
          <ControlTemplate TargetType="ListViewItem">
            <Border x:Name="Bd" Background="{TemplateBinding Background}"
                    BorderBrush="#F2F2F2" BorderThickness="0,0,0,1">
              <ContentPresenter/>
            </Border>
            <ControlTemplate.Triggers>
              <Trigger Property="IsMouseOver" Value="True">
                <Setter TargetName="Bd" Property="Background" Value="#F5F5FA"/>
              </Trigger>
              <Trigger Property="IsSelected" Value="True">
                <Setter TargetName="Bd" Property="Background" Value="#E8F1FF"/>
              </Trigger>
            </ControlTemplate.Triggers>
          </ControlTemplate>
        </Setter.Value>
      </Setter>
    </Style>

    <!-- Credit / link button -->
    <Style x:Key="CreditBtn" TargetType="Button">
      <Setter Property="Background"      Value="Transparent"/>
      <Setter Property="Foreground"      Value="#8E8E93"/>
      <Setter Property="FontSize"        Value="10"/>
      <Setter Property="BorderThickness" Value="0"/>
      <Setter Property="Cursor"          Value="Hand"/>
      <Setter Property="Padding"         Value="0"/>
      <Setter Property="Template">
        <Setter.Value>
          <ControlTemplate TargetType="Button">
            <TextBlock x:Name="Tb" Text="{TemplateBinding Content}"
                       Foreground="{TemplateBinding Foreground}"
                       FontSize="{TemplateBinding FontSize}"
                       VerticalAlignment="Center"/>
            <ControlTemplate.Triggers>
              <Trigger Property="IsMouseOver" Value="True">
                <Setter TargetName="Tb" Property="Foreground"        Value="#007AFF"/>
                <Setter TargetName="Tb" Property="TextDecorations"   Value="Underline"/>
              </Trigger>
            </ControlTemplate.Triggers>
          </ControlTemplate>
        </Setter.Value>
      </Setter>
    </Style>

  </Window.Resources>

  <Grid>
    <Grid.RowDefinitions>
      <RowDefinition Height="64"/>
      <RowDefinition Height="*"/>
      <RowDefinition Height="44"/>
    </Grid.RowDefinitions>

    <!-- ================== HEADER ================== -->
    <Border Grid.Row="0" Background="White" Panel.ZIndex="1">
      <Border.Effect>
        <DropShadowEffect BlurRadius="8" ShadowDepth="2" Direction="270"
                          Color="#000000" Opacity="0.07"/>
      </Border.Effect>
      <Grid Margin="20,0">
        <Grid.ColumnDefinitions>
          <ColumnDefinition Width="Auto"/>
          <ColumnDefinition Width="*"/>
          <ColumnDefinition Width="Auto"/>
          <ColumnDefinition Width="Auto"/>
          <ColumnDefinition Width="Auto"/>
        </Grid.ColumnDefinitions>

        <!-- App title -->
        <StackPanel Grid.Column="0" VerticalAlignment="Center" Margin="0,0,28,0">
          <TextBlock Text="SPC Dashboard" FontSize="18" FontWeight="SemiBold"
                     Foreground="#1C1C1E"/>
          <TextBlock x:Name="SubTitle" Text="Loading..." FontSize="11"
                     Foreground="#8E8E93"/>
        </StackPanel>

        <!-- Search -->
        <Grid Grid.Column="2" VerticalAlignment="Center" Margin="0,0,10,0" Width="240">
          <TextBox x:Name="SearchBox" Style="{StaticResource SearchStyle}"/>
          <TextBlock x:Name="SearchHint" Text="Search name or hostname..."
                     IsHitTestVisible="False" VerticalAlignment="Center"
                     Margin="12,0" Foreground="#AEAEB2" FontSize="12"/>
        </Grid>

        <!-- Sort -->
        <ComboBox x:Name="SortCombo" Grid.Column="3" VerticalAlignment="Center"
                  Width="130" Margin="0,0,10,0" FontSize="12" Padding="8,6">
          <ComboBoxItem Content="Sort: Name"    Tag="Name"    IsSelected="True"/>
          <ComboBoxItem Content="Sort: Status"  Tag="Status"/>
          <ComboBoxItem Content="Sort: Latency" Tag="Latency"/>
        </ComboBox>

        <!-- Refresh -->
        <Button x:Name="RefreshBtn" Grid.Column="4" Style="{StaticResource SecondaryBtn}"
                Content="Refresh" VerticalAlignment="Center"/>
      </Grid>
    </Border>

    <!-- ================== CONTENT ================== -->
    <Grid Grid.Row="1" Margin="14,14,14,6">
      <Grid.ColumnDefinitions>
        <ColumnDefinition Width="2*"/>
        <ColumnDefinition Width="10"/>
        <ColumnDefinition Width="360"/>
      </Grid.ColumnDefinitions>

      <!-- Server list card -->
      <Border Grid.Column="0" Background="White" CornerRadius="12">
        <Border.Effect>
          <DropShadowEffect BlurRadius="14" ShadowDepth="2" Opacity="0.07"/>
        </Border.Effect>
        <ListView x:Name="ServerList"
                  Background="Transparent" BorderThickness="0"
                  ItemContainerStyle="{StaticResource ListItem}"
                  ScrollViewer.HorizontalScrollBarVisibility="Disabled"
                  VirtualizingPanel.IsVirtualizing="True"
                  VirtualizingPanel.VirtualizationMode="Recycling">
          <ListView.ItemTemplate>
            <DataTemplate>
              <Grid Height="50" Margin="16,0,16,0">
                <Grid.ColumnDefinitions>
                  <ColumnDefinition Width="18"/>
                  <ColumnDefinition Width="*"/>
                  <ColumnDefinition Width="72"/>
                </Grid.ColumnDefinitions>

                <!-- Status dot -->
                <Ellipse Grid.Column="0" Width="10" Height="10" VerticalAlignment="Center">
                  <Ellipse.Style>
                    <Style TargetType="Ellipse">
                      <Setter Property="Fill" Value="#8E8E93"/>
                      <Style.Triggers>
                        <DataTrigger Binding="{Binding Status}" Value="online">
                          <Setter Property="Fill" Value="#34C759"/>
                        </DataTrigger>
                        <DataTrigger Binding="{Binding Status}" Value="offline">
                          <Setter Property="Fill" Value="#FF3B30"/>
                        </DataTrigger>
                      </Style.Triggers>
                    </Style>
                  </Ellipse.Style>
                </Ellipse>

                <!-- Name + hostname -->
                <StackPanel Grid.Column="1" VerticalAlignment="Center" Margin="6,0,0,0">
                  <TextBlock Text="{Binding DisplayName}" FontSize="13"
                             FontWeight="SemiBold" Foreground="#1C1C1E"/>
                  <TextBlock Text="{Binding Hostname}" FontSize="11"
                             Foreground="#8E8E93" Margin="0,1,0,0"/>
                </StackPanel>

                <!-- Latency -->
                <TextBlock Grid.Column="2" Text="{Binding LatencyText}"
                           VerticalAlignment="Center" HorizontalAlignment="Right"
                           FontSize="11">
                  <TextBlock.Style>
                    <Style TargetType="TextBlock">
                      <Setter Property="Foreground" Value="#AEAEB2"/>
                      <Style.Triggers>
                        <DataTrigger Binding="{Binding Status}" Value="online">
                          <Setter Property="Foreground" Value="#34C759"/>
                        </DataTrigger>
                        <DataTrigger Binding="{Binding Status}" Value="offline">
                          <Setter Property="Foreground" Value="#FF3B30"/>
                        </DataTrigger>
                      </Style.Triggers>
                    </Style>
                  </TextBlock.Style>
                </TextBlock>
              </Grid>
            </DataTemplate>
          </ListView.ItemTemplate>
        </ListView>
      </Border>

      <!-- Detail panel card -->
      <Border Grid.Column="2" Background="White" CornerRadius="12">
        <Border.Effect>
          <DropShadowEffect BlurRadius="14" ShadowDepth="2" Opacity="0.07"/>
        </Border.Effect>
        <Grid>
          <Grid.RowDefinitions>
            <RowDefinition Height="*"/>
            <RowDefinition Height="Auto"/>
          </Grid.RowDefinitions>

          <ScrollViewer Grid.Row="0" VerticalScrollBarVisibility="Auto"
                        Padding="18,18,18,0">
            <StackPanel x:Name="DetailContent">
              <TextBlock Text="&lt;- Select a server" FontSize="14"
                         Foreground="#C7C7CC" HorizontalAlignment="Center"
                         Margin="0,80,0,0"/>
            </StackPanel>
          </ScrollViewer>

          <Border Grid.Row="1" Padding="18,12,18,18">
            <Grid>
              <Grid.ColumnDefinitions>
                <ColumnDefinition Width="Auto"/>
                <ColumnDefinition Width="10"/>
                <ColumnDefinition Width="*"/>
              </Grid.ColumnDefinitions>
              <Button x:Name="WakeBtn" Grid.Column="0" Style="{StaticResource SecondaryBtn}"
                      Content="Wake" MinWidth="88" IsEnabled="False"/>
              <Button x:Name="ConnectBtn" Grid.Column="2" Style="{StaticResource PrimaryBtn}"
                      Content="Connect via VNC"
                      HorizontalAlignment="Stretch"
                      IsEnabled="False"/>
            </Grid>
          </Border>
        </Grid>
      </Border>
    </Grid>

    <!-- ================== FOOTER ================== -->
    <Border Grid.Row="2" Background="White" Margin="0,4,0,0">
      <Border.Effect>
        <DropShadowEffect BlurRadius="6" ShadowDepth="-1" Direction="90"
                          Color="#000000" Opacity="0.05"/>
      </Border.Effect>
      <Grid Margin="20,0">
        <Grid.ColumnDefinitions>
          <ColumnDefinition Width="*"/>
          <ColumnDefinition Width="Auto"/>
          <ColumnDefinition Width="Auto"/>
        </Grid.ColumnDefinitions>

        <TextBlock x:Name="StatusBar" Grid.Column="0" VerticalAlignment="Center"
                   FontSize="12" Foreground="#6E6E73" Text="Loading servers..."/>

        <TextBlock x:Name="AutoRefreshLbl" Grid.Column="1" VerticalAlignment="Center"
                   FontSize="11" Foreground="#AEAEB2" Margin="0,0,22,0"/>

        <Button x:Name="CreditBtn" Grid.Column="2" Style="{StaticResource CreditBtn}"
                Content="credits: ssari9@ford.com  contact: ithelp@ford.com"
                VerticalAlignment="Center"/>
      </Grid>
    </Border>

  </Grid>
</Window>
'@

# =============================== LOAD WINDOW =================================
$reader = New-Object System.Xml.XmlNodeReader($xaml)
$window = [System.Windows.Markup.XamlReader]::Load($reader)

# Grab named controls
$subTitle      = $window.FindName('SubTitle')
$searchBox     = $window.FindName('SearchBox')
$searchHint    = $window.FindName('SearchHint')
$sortCombo     = $window.FindName('SortCombo')
$refreshBtn    = $window.FindName('RefreshBtn')
$serverList    = $window.FindName('ServerList')
$detailContent = $window.FindName('DetailContent')
$wakeBtn       = $window.FindName('WakeBtn')
$connectBtn    = $window.FindName('ConnectBtn')
$statusBar     = $window.FindName('StatusBar')
$autoRefLbl    = $window.FindName('AutoRefreshLbl')
$creditBtn     = $window.FindName('CreditBtn')

# Observable collection + collection view
$script:ServerCollection = New-Object System.Collections.ObjectModel.ObservableCollection[object]
$serverList.ItemsSource  = $script:ServerCollection
$script:CollView         = [System.Windows.Data.CollectionViewSource]::GetDefaultView($script:ServerCollection)

# Fast lookup: Hostname -> ServerItem list
$script:ServerLookup = @{}

# Track selected item for detail panel refresh
$script:SelectedItem        = $null
$script:DetailStatusDot     = $null
$script:DetailStatusText    = $null
$script:StatusMessage       = $null
$script:StatusMessageTimer  = New-Object System.Windows.Threading.DispatcherTimer
$script:StatusMessageTimer.Interval = [TimeSpan]::FromSeconds(5)
$script:StatusMessageTimer.Add_Tick({
    $script:StatusMessageTimer.Stop()
    $script:StatusMessage = $null
    Update-StatusBar
})

# =================== HELPER: colour string -> Brush ===========================
function New-Brush([string]$hex) {
    $c = [System.Windows.Media.ColorConverter]::ConvertFromString($hex)
    return [System.Windows.Media.SolidColorBrush]::new($c)
}

function Get-HostLookupKey([string]$Hostname) {
    if ([string]::IsNullOrWhiteSpace($Hostname)) { return $null }
    return $Hostname.Trim().ToLowerInvariant()
}

function Get-ItemMacAddress($item) {
    if (-not $item -or -not $item.RawData) { return $null }
    if (-not $item.RawData.Contains($script:MacAddressColumnIndex)) { return $null }

    $value = [string]$item.RawData[$script:MacAddressColumnIndex]
    if ([string]::IsNullOrWhiteSpace($value)) { return $null }
    return $value.Trim()
}

function Update-ActionButtons($item) {
    $connectBtn.IsEnabled = [bool]($item -and $item.Hostname)
    $wakeBtn.IsEnabled = [bool]($item -and $item.Status -eq 'offline' -and (Get-ItemMacAddress $item))
}

function Show-StatusMessage([string]$message) {
    $script:StatusMessage = $message
    Update-StatusBar
    $script:StatusMessageTimer.Stop()
    $script:StatusMessageTimer.Start()
}

function ConvertTo-MacBytes([string]$MacAddress) {
    if ([string]::IsNullOrWhiteSpace($MacAddress)) { return $null }

    $hex = ($MacAddress -replace '[^0-9A-Fa-f]', '').ToUpperInvariant()
    if ($hex.Length -ne 12) { return $null }

    $bytes = New-Object byte[] 6
    for ($i = 0; $i -lt 6; $i++) {
        $bytes[$i] = [Convert]::ToByte($hex.Substring($i * 2, 2), 16)
    }
    return $bytes
}

function Send-WakeOnLan([string]$MacAddress) {
    $macBytes = ConvertTo-MacBytes $MacAddress
    if (-not $macBytes) {
        throw "Invalid MAC address: $MacAddress"
    }

    $packet = New-Object byte[] 102
    for ($i = 0; $i -lt 6; $i++) {
        $packet[$i] = 0xFF
    }
    for ($i = 6; $i -lt $packet.Length; $i += 6) {
        [System.Buffer]::BlockCopy($macBytes, 0, $packet, $i, 6)
    }

    $client = New-Object System.Net.Sockets.UdpClient
    try {
        $client.EnableBroadcast = $true
        [void]$client.Send($packet, $packet.Length, '255.255.255.255', 9)
    } finally {
        try { $client.Close() } catch {}
        try { $client.Dispose() } catch {}
    }
}

function Invoke-WakeServer($item) {
    if (-not $item) { return }

    $macAddress = Get-ItemMacAddress $item
    if (-not $macAddress) {
        Show-StatusMessage "Wake-on-LAN unavailable: no MAC address configured."
        return
    }

    try {
        Send-WakeOnLan $macAddress
        $targetName = if ($item.Hostname) { $item.Hostname } else { $item.DisplayName }
        Show-StatusMessage "WOL packet sent to $targetName."
    } catch {
        Show-StatusMessage "Wake-on-LAN failed: $($_.Exception.Message)"
    }
}

# ============================ DETAIL PANEL ===================================
function Update-DetailPanel {
    param($item)

    $detailContent.Children.Clear()
    $script:DetailStatusDot  = $null
    $script:DetailStatusText = $null
    $script:SelectedItem     = $item

    if (-not $item) {
        $ph = New-Object System.Windows.Controls.TextBlock
        $ph.Text = "<- Select a server"
        $ph.FontSize = 14
        $ph.Foreground = New-Brush '#C7C7CC'
        $ph.HorizontalAlignment = 'Center'
        $ph.Margin = [System.Windows.Thickness]::new(0,80,0,0)
        [void]$detailContent.Children.Add($ph)
        Update-ActionButtons $null
        return
    }

    # Title
    $titleTb = New-Object System.Windows.Controls.TextBlock
    $titleTb.Text = $item.DisplayName
    $titleTb.FontSize = 20
    $titleTb.FontWeight = [System.Windows.FontWeights]::Bold
    $titleTb.Foreground = New-Brush '#1C1C1E'
    $titleTb.Margin = [System.Windows.Thickness]::new(0,0,0,6)
    $titleTb.TextWrapping = 'Wrap'
    [void]$detailContent.Children.Add($titleTb)

    # Status row
    $statusRow = New-Object System.Windows.Controls.StackPanel
    $statusRow.Orientation = 'Horizontal'
    $statusRow.Margin = [System.Windows.Thickness]::new(0,0,0,14)

    $dot = New-Object System.Windows.Shapes.Ellipse
    $dot.Width  = 10
    $dot.Height = 10
    $dot.VerticalAlignment = 'Center'
    $dot.Margin = [System.Windows.Thickness]::new(0,0,6,0)
    $dot.Fill = New-Brush (Get-StatusColor $item.Status)
    $script:DetailStatusDot = $dot

    $stTb = New-Object System.Windows.Controls.TextBlock
    $stTb.Text = Get-StatusText $item
    $stTb.FontSize = 13
    $stTb.Foreground = New-Brush (Get-StatusColor $item.Status)
    $stTb.VerticalAlignment = 'Center'
    $script:DetailStatusText = $stTb

    [void]$statusRow.Children.Add($dot)
    [void]$statusRow.Children.Add($stTb)
    [void]$detailContent.Children.Add($statusRow)

    # Separator
    $sep = New-Object System.Windows.Controls.Separator
    $sep.Margin = [System.Windows.Thickness]::new(0,0,0,12)
    [void]$detailContent.Children.Add($sep)

    # Column values
    foreach ($cd in $script:ColumnDefs) {
        $val = if ($item.RawData.Contains($cd.Index)) { $item.RawData[$cd.Index] } else { '' }
        if (-not $val) { continue }
        Add-DetailRow $cd.Label $val
    }

    Update-ActionButtons $item
}

function Get-StatusColor([string]$status) {
    switch ($status) {
        'online'  { return '#34C759' }
        'offline' { return '#FF3B30' }
        default   { return '#8E8E93' }
    }
}

function Get-StatusText($item) {
    switch ($item.Status) {
        'online'  { return "Online - $($item.LatencyMs) ms" }
        'offline' { return 'Offline' }
        default   { return 'Checking...' }
    }
}

function Add-DetailRow([string]$label, [string]$value) {
    $g   = New-Object System.Windows.Controls.Grid
    $g.Margin = [System.Windows.Thickness]::new(0,0,0,8)
    $c1  = New-Object System.Windows.Controls.ColumnDefinition
    $c1.Width = [System.Windows.GridLength]::new(130)
    $c2  = New-Object System.Windows.Controls.ColumnDefinition
    $c2.Width = [System.Windows.GridLength]::new(1,[System.Windows.GridUnitType]::Star)
    [void]$g.ColumnDefinitions.Add($c1)
    [void]$g.ColumnDefinitions.Add($c2)

    $lbl = New-Object System.Windows.Controls.TextBlock
    $lbl.Text = $label
    $lbl.FontSize = 11
    $lbl.Foreground = New-Brush '#8E8E93'
    $lbl.VerticalAlignment = 'Top'
    $lbl.TextWrapping = 'Wrap'
    [System.Windows.Controls.Grid]::SetColumn($lbl, 0)

    $val = New-Object System.Windows.Controls.TextBlock
    $val.Text = $value
    $val.FontSize = 12
    $val.Foreground = New-Brush '#1C1C1E'
    $val.VerticalAlignment = 'Top'
    $val.TextWrapping = 'Wrap'
    [System.Windows.Controls.Grid]::SetColumn($val, 1)

    [void]$g.Children.Add($lbl)
    [void]$g.Children.Add($val)
    [void]$detailContent.Children.Add($g)
}

# Refresh only the status portion of the detail panel (called from ping timer)
function Refresh-DetailStatus {
    $item = $script:SelectedItem
    if (-not $item -or -not $script:DetailStatusDot) { return }
    $color = New-Brush (Get-StatusColor $item.Status)
    $script:DetailStatusDot.Fill  = $color
    $script:DetailStatusText.Text = Get-StatusText $item
    $script:DetailStatusText.Foreground = $color
    Update-ActionButtons $item
}

# =============================== STATUS BAR ==================================
function Update-StatusBar {
    $total    = $script:ServerCollection.Count
    $online   = ($script:ServerCollection | Where-Object Status -eq 'online').Count
    $offline  = ($script:ServerCollection | Where-Object Status -eq 'offline').Count
    $checking = ($script:ServerCollection | Where-Object Status -eq 'checking').Count

    $text = "$total servers - "
    $text += "$online online - $offline offline"
    if ($checking -gt 0) { $text += " - $checking checking..." }
    if ($script:StatusMessage) { $text += " | $script:StatusMessage" }
    $statusBar.Text = $text
    $subTitle.Text  = "$online / $total online"
}

# ============================== SORT / FILTER =================================
function Apply-Sort {
    $script:CollView.SortDescriptions.Clear()
    $selectedSortItem = $sortCombo.SelectedItem -as [System.Windows.Controls.ComboBoxItem]
    $tag = if ($selectedSortItem) { $selectedSortItem.Tag } else { $null }
    if (-not $tag) { $tag = 'Name' }
    switch ($tag) {
        'Name'    { $script:CollView.SortDescriptions.Add([System.ComponentModel.SortDescription]::new('DisplayName',[System.ComponentModel.ListSortDirection]::Ascending)) }
        'Status'  { $script:CollView.SortDescriptions.Add([System.ComponentModel.SortDescription]::new('Status',     [System.ComponentModel.ListSortDirection]::Ascending)) }
        'Latency' { $script:CollView.SortDescriptions.Add([System.ComponentModel.SortDescription]::new('LatencyMs',  [System.ComponentModel.ListSortDirection]::Ascending)) }
    }
}

$script:CurrentFilter = ''
function Apply-Filter {
    $ft = $searchBox.Text.Trim().ToLower()
    $script:CurrentFilter = $ft
    if ($ft) {
        $script:CollView.Filter = {
            param($obj)
            $s = [ServerItem]$obj
            ($s.DisplayName -and $s.DisplayName.ToLower().Contains($script:CurrentFilter)) -or
            ($s.Hostname    -and $s.Hostname.ToLower().Contains($script:CurrentFilter))
        }
    } else {
        $script:CollView.Filter = $null
    }
}

# ============================== LOAD / RELOAD =================================
function Load-Servers {
    $statusBar.Text = "Reading Excel file..."
    $excelPath = Resolve-ExcelFilePath
    $items = Read-ExcelServers -Path $excelPath
    $script:ServerCollection.Clear()
    $script:ServerLookup.Clear()
    $prevSelected = $null
    foreach ($item in $items) {
        $script:ServerCollection.Add($item) | Out-Null
        $key = Get-HostLookupKey $item.Hostname
        if ($key) {
            if (-not $script:ServerLookup.ContainsKey($key)) {
                $script:ServerLookup[$key] = [System.Collections.Generic.List[object]]::new()
            }
            $script:ServerLookup[$key].Add($item) | Out-Null
        }
    }
    Apply-Sort
    Apply-Filter
    Update-StatusBar
    $subTitle.Text = "$($script:ServerCollection.Count) servers loaded"
    return $items
}

function Refresh-All {
    $refreshBtn.IsEnabled = $false
    $refreshBtn.Content   = "Refreshing..."
    try {
        $items = Load-Servers
        if ($items.Count -gt 0) {
            Start-PingAll -Collection $script:ServerCollection
        } else {
            $statusBar.Text = "No servers loaded. Check Excel file path."
        }
    } finally {
        $refreshBtn.IsEnabled = $true
        $refreshBtn.Content   = "Refresh"
    }
}

# ============================== VNC CONNECT ==================================
function Connect-VNC {
    param($item)
    if (-not $item -or -not $item.Hostname) {
        [System.Windows.MessageBox]::Show(
            "This server has no hostname configured.",
            "Cannot Connect", "OK", "Warning") | Out-Null
        return
    }
    if (-not (Test-Path $script:Config.VncExe)) {
        [System.Windows.MessageBox]::Show(
            "UltraVNC not found at:`n$($script:Config.VncExe)`n`nUpdate VncExe in SPC-Dashboard.ps1.",
            "VNC Not Found", "OK", "Warning") | Out-Null
        return
    }
    $pw   = $script:Config.VncPassword
    $host = $item.Hostname
    Start-Process -FilePath $script:Config.VncExe -ArgumentList "-password `"$pw`" -connect $host"
}

# ============================== EVENT HANDLERS ===============================

# Search box
$searchBox.Add_TextChanged({
    $searchHint.Visibility = if ($searchBox.Text) { 'Collapsed' } else { 'Visible' }
    Apply-Filter
})

# Sort
$sortCombo.Add_SelectionChanged({ Apply-Sort })

# Refresh button
$refreshBtn.Add_Click({ Refresh-All })

# List selection -> update detail panel
$serverList.Add_SelectionChanged({
    Update-DetailPanel ($serverList.SelectedItem -as [ServerItem])
})

# Double-click -> connect
$serverList.Add_MouseDoubleClick({
    $item = $serverList.SelectedItem -as [ServerItem]
    if ($item) { Connect-VNC $item }
})

# Connect button
$connectBtn.Add_Click({
    Connect-VNC ($serverList.SelectedItem -as [ServerItem])
})

# Wake button
$wakeBtn.Add_Click({
    Invoke-WakeServer ($serverList.SelectedItem -as [ServerItem])
})

# Credit button -> open default mail client
$creditBtn.Add_Click({
    Start-Process "mailto:ithelp@ford.com"
})

# ================================ TIMERS =====================================

# Poll ping results every 200 ms and update UI
$pingTimer = New-Object System.Windows.Threading.DispatcherTimer
$pingTimer.Interval = [TimeSpan]::FromMilliseconds(200)
$pingTimer.Add_Tick({
    $result = $null
    $updated = $false
    while ($script:PingQueue -and $script:PingQueue.TryDequeue([ref]$result)) {
        $matchedItems = $script:ServerLookup[(Get-HostLookupKey $result.Host)]
        if ($matchedItems) {
            foreach ($item in $matchedItems) {
                $item.Status    = $result.Status
                $item.LatencyMs = $result.Latency
                $updated = $true
                # Refresh detail panel status if this is the selected server
                if ($script:SelectedItem -eq $item) {
                    Refresh-DetailStatus
                }
            }
        }
    }
    if ($updated) { Update-StatusBar }
})
$pingTimer.Start()

# Auto-refresh countdown
$script:AutoRefCountdown = $script:Config.AutoRefreshSeconds
$autoRefTimer = New-Object System.Windows.Threading.DispatcherTimer
if ($script:Config.AutoRefreshSeconds -gt 0) {
    $autoRefTimer.Interval = [TimeSpan]::FromSeconds(1)
    $autoRefTimer.Add_Tick({
        $script:AutoRefCountdown--
        if ($script:AutoRefCountdown -le 0) {
            $script:AutoRefCountdown = $script:Config.AutoRefreshSeconds
            Refresh-All
        }
        $autoRefLbl.Text = "Auto-refresh in $script:AutoRefCountdown s"
    })
    $autoRefTimer.Start()
    $autoRefLbl.Text = "Auto-refresh in $script:AutoRefCountdown s"
} else {
    $autoRefLbl.Text = "Auto-refresh: off"
}

# ================= FILE WATCHER (auto-reload when Excel changes) ==============
$script:WatcherRegistered = $false
function Register-FileWatcher {
    $excelPath = Resolve-ExcelFilePath
    if (-not (Test-Path $excelPath)) { return }

    $dir  = [IO.Path]::GetDirectoryName($excelPath)
    $file = [IO.Path]::GetFileName($excelPath)

    $script:Watcher = New-Object IO.FileSystemWatcher($dir, $file)
    $script:Watcher.NotifyFilter = [IO.NotifyFilters]::LastWrite
    $script:Watcher.EnableRaisingEvents = $true

    $script:WatcherJob = Register-ObjectEvent -InputObject $script:Watcher `
        -EventName Changed -SourceIdentifier 'SPC_ExcelChanged' -Action {
        Start-Sleep -Seconds 2   # wait for Excel to finish writing
        $window.Dispatcher.Invoke([action]{
            $statusBar.Text = "Excel file changed - reloading..."
            $items = Load-Servers
            if ($items.Count -gt 0) { Start-PingAll -Collection $script:ServerCollection }
        })
    }
    $script:WatcherRegistered = $true
}

# ============================= WINDOW EVENTS ==================================
$window.Add_Loaded({
    Refresh-All
    Register-FileWatcher
})

$window.Add_Closed({
    $pingTimer.Stop()
    $autoRefTimer.Stop()
    if ($script:WatcherRegistered) {
        Unregister-Event -SourceIdentifier 'SPC_ExcelChanged' -ErrorAction SilentlyContinue
        $script:Watcher.Dispose()
    }
    if ($script:PingRunspace) {
        try { $script:PingRunspace.Close(); $script:PingRunspace.Dispose() } catch {}
    }
    if ($script:PingPS) {
        try { $script:PingPS.Dispose() } catch {}
    }
})

# ================================= SHOW ======================================
[void]$window.ShowDialog()
