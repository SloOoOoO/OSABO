# SPC Server Dashboard

A standalone, zero-dependency Windows dashboard for monitoring SPC measurement-station servers.  
Double-click a server to connect via UltraVNC.  No Python, Node.js, or additional software required — just PowerShell 5.1+ (built-in since Windows 10).

---

## Quick Start

1. **Clone / download** this repository so `SPC-Dashboard.ps1`, `Start-Dashboard.bat`, and  
   `SPC-Gerätenliste_Gesamt_NEU.xlsm` are all in the **same folder**.
2. **Set your VNC password** — see [Configuration](#configuration) below.  
   ⚠️  Never commit a real password to the repository.
3. **Double-click `Start-Dashboard.bat`** — the dashboard opens automatically.

---

## Requirements

| Requirement | Detail |
|---|---|
| Windows OS | Windows 10 / 11 (or Windows 7/8 with .NET 4.5+) |
| PowerShell | 5.1 (built-in) — no extra install needed |
| .NET Framework | 4.5 or newer (already present on all supported Windows) |
| UltraVNC | `C:\Legacy\UltraVNC\UltraVNC-Viewer.exe` (configurable) |
| Excel | **Not required** — the file is parsed as ZIP+XML natively |

---

## Configuration

Open `SPC-Dashboard.ps1` in any text editor and edit the `$script:Config` block near the top:

```powershell
$script:Config = @{
    # Path to the Excel workbook (defaults to same folder as this script)
    ExcelFile          = Join-Path $PSScriptRoot "SPC-Gerätenliste_Gesamt_NEU.xlsm"

    # Sheet index (1-based). 1 = "Handmessplätze 09.08.2023"
    SheetIndex         = 1

    # First data row (rows 1–2 are headers)
    DataStartRow       = 3

    # Full path to the UltraVNC viewer executable
    VncExe             = "C:\Legacy\UltraVNC\UltraVNC-Viewer.exe"

    # VNC password — set here OR via the VNC_PASSWORD environment variable
    VncPassword        = "CHANGE_ME"    # <-- replace with real password

    # How often to automatically re-ping all servers (seconds). 0 = disabled.
    AutoRefreshSeconds = 30

    # Ping timeout per host (milliseconds)
    PingTimeoutMs      = 1500
}
```

### Setting the password via environment variable (recommended)

Instead of editing the script, set the `VNC_PASSWORD` environment variable before launching:

```batch
:: In Start-Dashboard.bat, uncomment and edit this line:
set VNC_PASSWORD=YourRealPasswordHere
```

The script reads `$env:VNC_PASSWORD` at startup and prefers it over the hardcoded value.

---

## Usage

| Action | How |
|---|---|
| **View server list** | All SPC servers from the Excel file are listed with live online-status indicators |
| **Filter** | Type in the Search box — filters by SPC name or hostname in real time |
| **Sort** | Use the Sort drop-down in the header (Name / Status / Latency) |
| **See details** | Click any server — all Excel columns appear in the right-hand panel |
| **Connect via VNC** | Double-click any server **or** select it and click **Connect via VNC** |
| **Manual refresh** | Click **↺ Refresh** in the header |
| **Auto-refresh** | Pings all servers every 30 s (configurable); countdown shown in footer |
| **Excel hot-reload** | When the `.xlsm` file is saved/updated, the list reloads automatically |

### Status indicators

| Colour | Meaning |
|---|---|
| 🟢 Green | Online — latency shown in ms |
| 🔴 Red | Offline / unreachable |
| ⚫ Gray | Checking (initial ping in progress) |

---

## Excel File Structure

**File:** `SPC-Gerätenliste_Gesamt_NEU.xlsm`  
**Active sheet:** Sheet 1 — *"Handmessplätze 09.08.2023"*  
**Data range:** Rows 3–175 (≈ 160 servers); rows 1–2 are header rows.

| Column | Index | Header label | Role in dashboard |
|---|---|---|---|
| A | 0 | SPC-Nr. | **Display name** shown in the server list |
| B | 1 | Typ | Detail panel |
| C | 2 | Teil-Name / Teil-Nummer | Detail panel |
| D | 3 | Op. | Detail panel |
| E | 4 | Modell | Detail panel |
| F | 5 | Standort / Halle | Detail panel |
| G | 6 | Pfeiler | Detail panel |
| H | 7 | Tisch-Nr. | Detail panel |
| I | 8 | PTM | Detail panel |
| J | 9 | Serien-Nr. | Detail panel |
| K | 10 | Prüfplan / Messprog. Version | Detail panel |
| L | 11 | Kanal-Belegung Anzahl | Detail panel |
| M | 12 | Kanal belegt | Detail panel |
| N | 13 | Sonder-HW | Detail panel |
| O | 14 | Anwahlbox | Detail panel |
| P | 15 | Windows-Version | Detail panel |
| Q | 16 | IR-57 Sichtprüfung | Detail panel |
| R | 17 | Filter / Lüfter | Detail panel |
| S | 18 | NW-fähig | Detail panel |
| T | 19 | Anzahl Prüfplan | Detail panel |
| U | 20 | Daten gesichert | Detail panel |
| V | 21 | Daten upload | Detail panel |
| W | 22 | Datenmanager umgestellt | Detail panel |
| X | 23 | Greengate | Detail panel |
| Y | 24 | Bemerkung | Detail panel |
| **Z** | **25** | **Hostname** | **VNC connect target** (e.g. `SPC48478`) |
| AA | 26 | MAC-Adresse | Detail panel |
| AB | 27 | Serien-Nr. USB | Detail panel |
| AC | 28 | Green Gate korrekt | Detail panel |

### Adding new servers

Simply add a new row to the Excel sheet with at minimum:
- **Column A** — the display name (e.g. `G-200`)
- **Column Z** — the hostname or IP address

Save the file. If the dashboard is already open, it reloads automatically within ~2 seconds.

---

## Troubleshooting

| Problem | Solution |
|---|---|
| "Excel file not found" error | Ensure the `.xlsm` file is in the same folder as `SPC-Dashboard.ps1`, or update `ExcelFile` in the config |
| "VNC Not Found" dialog | Update `VncExe` in the config to the correct path |
| All servers show offline | Check network connectivity; adjust `PingTimeoutMs` if the network is slow |
| Script won't run | Right-click `Start-Dashboard.bat` → *Run as administrator* |
| Script blocked by policy | The `.bat` launcher already passes `-ExecutionPolicy Bypass`; if blocked further, check Group Policy |

---

## Credits

- **Developer:** ssari9@ford.com  
- **Support / IT Help:** ithelp@ford.com