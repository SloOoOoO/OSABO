@echo off
:: SPC Server Dashboard Launcher
:: Starts the PowerShell WPF dashboard without requiring a changed execution policy system-wide.

title SPC Server Dashboard

:: Optionally set the VNC password via environment variable (more secure than editing the PS1):
:: set VNC_PASSWORD=YourPasswordHere

powershell.exe -ExecutionPolicy Bypass -NoProfile -WindowStyle Hidden -File "%~dp0SPC-Dashboard.ps1"

:: If the script exited with an error, pause so you can read the message.
if %ERRORLEVEL% NEQ 0 (
    echo.
    echo Dashboard exited with error code %ERRORLEVEL%.
    pause
)
