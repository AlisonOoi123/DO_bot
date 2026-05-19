# DO Bot — Windows Production Setup
# Run this script ONCE in PowerShell as Administrator
# Usage: Right-click PowerShell → "Run as Administrator", then:
#   cd "C:\Users\Administrator\OneDrive\Documents\GitHub\DO_bot"
#   .\setup_windows.ps1

$ErrorActionPreference = "Stop"
$APP_DIR = "C:\Users\Administrator\OneDrive\Documents\GitHub\DO_bot"
$NSSM    = "$APP_DIR\nssm.exe"
$PYTHON  = (Get-Command python).Source

Set-Location $APP_DIR

# ── Step 1: Install Python dependencies ──────────────────────────────────────
Write-Host "`n=== [1/4] Installing Python dependencies ===" -ForegroundColor Cyan
python -m pip install -r requirements.txt

# ── Step 2: Download NSSM (service manager) ───────────────────────────────────
Write-Host "`n=== [2/4] Downloading NSSM ===" -ForegroundColor Cyan
if (-not (Test-Path $NSSM)) {
    $nssmZip = "$env:TEMP\nssm.zip"
    Invoke-WebRequest -Uri "https://nssm.cc/release/nssm-2.24.zip" -OutFile $nssmZip
    Expand-Archive -Path $nssmZip -DestinationPath "$env:TEMP\nssm_extract" -Force
    Copy-Item "$env:TEMP\nssm_extract\nssm-2.24\win64\nssm.exe" $NSSM
    Write-Host "   nssm.exe saved to $NSSM"
} else {
    Write-Host "   nssm.exe already present, skipping."
}

# ── Step 3: Install DO Bot as Windows Service ─────────────────────────────────
Write-Host "`n=== [3/4] Installing DO Bot service ===" -ForegroundColor Cyan
& $NSSM stop  do_bot 2>$null
& $NSSM remove do_bot confirm 2>$null

& $NSSM install do_bot $PYTHON "$APP_DIR\app.py"
& $NSSM set     do_bot AppDirectory   $APP_DIR
& $NSSM set     do_bot DisplayName    "DO Bot - WhatsApp Lorry Assignment"
& $NSSM set     do_bot Description    "WhatsApp AI Logistics bot (Flask + waitress)"
& $NSSM set     do_bot Start          SERVICE_AUTO_START
& $NSSM set     do_bot AppStdout      "$APP_DIR\logs\do_bot.log"
& $NSSM set     do_bot AppStderr      "$APP_DIR\logs\do_bot_error.log"
& $NSSM set     do_bot AppRotateFiles 1
& $NSSM set     do_bot AppRotateBytes 5242880

# ── Step 4: Install ngrok as Windows Service ──────────────────────────────────
Write-Host "`n=== [4/4] Installing ngrok service ===" -ForegroundColor Cyan

$ngrokPath = (Get-Command ngrok -ErrorAction SilentlyContinue)?.Source
if (-not $ngrokPath) {
    Write-Host "   ngrok not found in PATH." -ForegroundColor Yellow
    Write-Host "   Download from https://ngrok.com/download, extract ngrok.exe to $APP_DIR" -ForegroundColor Yellow
    $ngrokPath = "$APP_DIR\ngrok.exe"
}

& $NSSM stop   ngrok_do_bot 2>$null
& $NSSM remove ngrok_do_bot confirm 2>$null

& $NSSM install ngrok_do_bot $ngrokPath "start do_bot --config `"$APP_DIR\ngrok_do_bot.yml`""
& $NSSM set     ngrok_do_bot AppDirectory $APP_DIR
& $NSSM set     ngrok_do_bot DisplayName  "DO Bot - ngrok Tunnel"
& $NSSM set     ngrok_do_bot Start        SERVICE_AUTO_START
& $NSSM set     ngrok_do_bot AppStdout    "$APP_DIR\logs\ngrok.log"
& $NSSM set     ngrok_do_bot AppStderr    "$APP_DIR\logs\ngrok_error.log"
& $NSSM set     ngrok_do_bot DependOnService do_bot

# ── Create logs directory ──────────────────────────────────────────────────────
New-Item -ItemType Directory -Force -Path "$APP_DIR\logs" | Out-Null

Write-Host "`n=== Setup complete! ===" -ForegroundColor Green
Write-Host ""
Write-Host "Before starting, make sure:" -ForegroundColor Yellow
Write-Host "  1. config.txt has your real META_ACCESS_TOKEN and META_PHONE_NUMBER_ID"
Write-Host "  2. ngrok_do_bot.yml has your real ngrok authtoken"
Write-Host "     (copy from ngrok_do_bot.yml.example and fill in the token)"
Write-Host "     Get your token at: https://dashboard.ngrok.com/get-started/your-authtoken"
Write-Host ""
Write-Host "Then start both services:" -ForegroundColor Cyan
Write-Host "  Start-Service do_bot"
Write-Host "  Start-Service ngrok_do_bot"
Write-Host ""
Write-Host "Check logs in: $APP_DIR\logs\"
