# DO Bot -- Windows Production Setup
# Run this script ONCE in PowerShell as Administrator
# Usage: Right-click PowerShell -> "Run as Administrator", then:
#   cd "C:\Users\Administrator\OneDrive\Documents\GitHub\DO_bot"
#   .\setup_windows.ps1

$ErrorActionPreference = "Stop"
$APP_DIR = "C:\Users\Administrator\OneDrive\Documents\GitHub\DO_bot"
$NSSM    = "$APP_DIR\nssm.exe"

# Use sys.executable to get the real python.exe, not the Windows Store alias
# (WindowsApps\python.exe is an app alias that the SYSTEM account cannot execute)
$PYTHON = python -c "import sys; print(sys.executable)"
if (-not $PYTHON -or -not (Test-Path $PYTHON)) {
    Write-Host "ERROR: Could not locate real python.exe. Got: $PYTHON" -ForegroundColor Red
    exit 1
}
Write-Host "Using Python: $PYTHON" -ForegroundColor Gray

# Derive the Python install root (two levels up from python.exe)
$PYTHON_DIR = Split-Path $PYTHON

Set-Location $APP_DIR

# -- Create logs directory FIRST (NSSM needs it before starting the service) --
New-Item -ItemType Directory -Force -Path "$APP_DIR\logs" | Out-Null
Write-Host "   Logs directory: $APP_DIR\logs" -ForegroundColor Gray

# -- Step 1: Install Python dependencies --------------------------------------
Write-Host "`n=== [1/4] Installing Python dependencies ===" -ForegroundColor Cyan
python -m pip install -r requirements.txt

# Resolve site-packages so it can be passed to the service environment
$SITE_PACKAGES = python -c "import site; print(site.getsitepackages()[0])" 2>$null
if (-not $SITE_PACKAGES) {
    $SITE_PACKAGES = python -c "import site; print(site.getusersitepackages())" 2>$null
}

# -- Step 2: Grant SYSTEM access to the user-scoped Python install ------------
# Services run as SYSTEM by default, which has no access to AppData\Local.
# icacls grants read+execute recursively without needing a password.
Write-Host "`n=== [2/4] Granting SYSTEM access to Python install ===" -ForegroundColor Cyan
$paths = @($PYTHON_DIR)
if ($SITE_PACKAGES -and (Test-Path $SITE_PACKAGES)) {
    $paths += $SITE_PACKAGES
}
foreach ($p in $paths) {
    Write-Host "   icacls: $p" -ForegroundColor Gray
    icacls $p /grant "NT AUTHORITY\SYSTEM:(OI)(CI)RX" /T /Q
}
# Also grant SYSTEM full access to the app directory (for writing logs, data files)
icacls $APP_DIR /grant "NT AUTHORITY\SYSTEM:(OI)(CI)F" /T /Q
Write-Host "   Done." -ForegroundColor Green

# -- Step 3: Download NSSM ----------------------------------------------------
Write-Host "`n=== [3/4] Downloading NSSM ===" -ForegroundColor Cyan
if (-not (Test-Path $NSSM)) {
    $nssmZip = "$env:TEMP\nssm.zip"
    Invoke-WebRequest -Uri "https://nssm.cc/release/nssm-2.24.zip" -OutFile $nssmZip
    Expand-Archive -Path $nssmZip -DestinationPath "$env:TEMP\nssm_extract" -Force
    Copy-Item "$env:TEMP\nssm_extract\nssm-2.24\win64\nssm.exe" $NSSM
    Write-Host "   nssm.exe saved to $NSSM"
} else {
    Write-Host "   nssm.exe already present, skipping."
}

# -- Step 4: Install services -------------------------------------------------
Write-Host "`n=== [4/4] Installing services ===" -ForegroundColor Cyan

# -- DO Bot service --
try { & $NSSM stop   do_bot         2>$null } catch {}; $LASTEXITCODE = 0
try { & $NSSM remove do_bot confirm 2>$null } catch {}; $LASTEXITCODE = 0

& $NSSM install do_bot $PYTHON
& $NSSM set     do_bot AppParameters    "$APP_DIR\app.py"
& $NSSM set     do_bot AppDirectory     $APP_DIR
& $NSSM set     do_bot DisplayName      "DO Bot - WhatsApp Lorry Assignment"
& $NSSM set     do_bot Description      "WhatsApp AI Logistics bot (Flask + waitress)"
& $NSSM set     do_bot Start            SERVICE_AUTO_START
& $NSSM set     do_bot AppStdout        "$APP_DIR\logs\do_bot.log"
& $NSSM set     do_bot AppStderr        "$APP_DIR\logs\do_bot_error.log"
& $NSSM set     do_bot AppRotateFiles   1
& $NSSM set     do_bot AppRotateBytes   5242880
& $NSSM set     do_bot AppNoConsole     1

$envExtra = "PYTHONUTF8=1"
if ($SITE_PACKAGES) {
    $envExtra = "PYTHONPATH=$SITE_PACKAGES`nPYTHONUTF8=1"
    Write-Host "   PYTHONPATH set to: $SITE_PACKAGES" -ForegroundColor Gray
}
& $NSSM set do_bot AppEnvironmentExtra $envExtra

# -- ngrok service --
# Search for ngrok in PATH, then common install locations (Scoop, Chocolatey, WinGet, etc.)
# WindowsApps aliases are excluded - they are Store stubs that SYSTEM cannot execute.
function Find-NgrokPath {
    $cmd = Get-Command ngrok -ErrorAction SilentlyContinue
    if ($cmd -and $cmd.Source -notlike "*WindowsApps*") { return $cmd.Source }

    $fixed = @(
        "$APP_DIR\ngrok.exe",
        "$env:USERPROFILE\scoop\shims\ngrok.exe",
        "$env:USERPROFILE\AppData\Local\ngrok\ngrok.exe",
        "$env:USERPROFILE\ngrok.exe",
        "C:\ProgramData\chocolatey\bin\ngrok.exe",
        "C:\Program Files\ngrok\ngrok.exe",
        "C:\ngrok\ngrok.exe"
    )
    foreach ($c in $fixed) {
        if (Test-Path $c) { return $c }
    }

    # WinGet installs into a versioned subdirectory - use wildcard
    $wingetMatch = Get-Item "$env:LOCALAPPDATA\Microsoft\WinGet\Packages\Ngrok.Ngrok*\ngrok.exe" `
        -ErrorAction SilentlyContinue | Select-Object -First 1 -ExpandProperty FullName
    if ($wingetMatch -and $wingetMatch -notlike "*WindowsApps*") { return $wingetMatch }

    return $null
}

function Install-NgrokStandalone {
    Write-Host "   Attempting to download standalone ngrok.exe directly to $APP_DIR ..." -ForegroundColor Cyan
    # Add a Defender exclusion for APP_DIR so the download isn't quarantined
    try {
        Add-MpPreference -ExclusionPath $APP_DIR -ErrorAction Stop
        Write-Host "   Windows Defender exclusion added for $APP_DIR" -ForegroundColor Gray
    } catch {
        Write-Host "   Could not add Defender exclusion (non-fatal): $_" -ForegroundColor Gray
    }

    # Download directly into APP_DIR (which has the Defender exclusion).
    # Do NOT use $env:TEMP -- Defender scans it and quarantines the file there
    # before it can be moved to the excluded folder.
    $zipPath = "$APP_DIR\ngrok-standalone.zip"
    try {
        Invoke-WebRequest -Uri "https://bin.equinox.io/c/bNyj1mQVY4c/ngrok-v3-stable-windows-amd64.zip" `
            -OutFile $zipPath -UseBasicParsing -ErrorAction Stop
        Expand-Archive -Path $zipPath -DestinationPath $APP_DIR -Force
        Remove-Item $zipPath -ErrorAction SilentlyContinue
        if (Test-Path "$APP_DIR\ngrok.exe") {
            Write-Host "   ngrok.exe extracted to $APP_DIR" -ForegroundColor Green
            return "$APP_DIR\ngrok.exe"
        }
    } catch {
        Write-Host "   Download failed: $_" -ForegroundColor Red
    }
    return $null
}

$ngrokPath = Find-NgrokPath
if (-not $ngrokPath) {
    Write-Host "   ngrok standalone not found -- downloading now..." -ForegroundColor Yellow
    $ngrokPath = Install-NgrokStandalone
    if (-not $ngrokPath) {
        $ngrokPath = "$APP_DIR\ngrok.exe"
        Write-Host ""
        Write-Host "   MANUAL STEP REQUIRED:" -ForegroundColor Red
        Write-Host "   1. Open Windows Security > Virus & threat protection > Exclusions" -ForegroundColor Red
        Write-Host "      and add an exclusion for: $APP_DIR" -ForegroundColor Red
        Write-Host "   2. Extract ngrok.exe from the zip at https://ngrok.com/download" -ForegroundColor Red
        Write-Host "      into $APP_DIR, then re-run this script." -ForegroundColor Red
    }
} else {
    Write-Host "   ngrok found: $ngrokPath" -ForegroundColor Gray
}

try { & $NSSM stop   ngrok_do_bot         2>$null } catch {}; $LASTEXITCODE = 0
try { & $NSSM remove ngrok_do_bot confirm 2>$null } catch {}; $LASTEXITCODE = 0

& $NSSM install ngrok_do_bot $ngrokPath
& $NSSM set     ngrok_do_bot AppParameters   "start do_bot --config `"$APP_DIR\ngrok_do_bot.yml`""
& $NSSM set     ngrok_do_bot AppDirectory    $APP_DIR
& $NSSM set     ngrok_do_bot DisplayName     "DO Bot - ngrok Tunnel"
& $NSSM set     ngrok_do_bot Start           SERVICE_AUTO_START
& $NSSM set     ngrok_do_bot AppStdout       "$APP_DIR\logs\ngrok.log"
& $NSSM set     ngrok_do_bot AppStderr       "$APP_DIR\logs\ngrok_error.log"
& $NSSM set     ngrok_do_bot AppNoConsole    1
& $NSSM set     ngrok_do_bot DependOnService do_bot

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
Write-Host ""
Write-Host "If services fail to start, run: .\diagnose_windows.ps1" -ForegroundColor Yellow
