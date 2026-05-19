# DO Bot -- Windows Service Diagnostics
# Run in PowerShell (Administrator recommended)
# Usage: .\diagnose_windows.ps1

$APP_DIR = "C:\Users\Administrator\OneDrive\Documents\GitHub\DO_bot"
$NSSM    = "$APP_DIR\nssm.exe"

Write-Host ""
Write-Host "=== DO Bot Service Diagnostics ===" -ForegroundColor Cyan

# -- 1. Check logs directory --------------------------------------------------
Write-Host ""
Write-Host "[1] Logs directory" -ForegroundColor Yellow
if (Test-Path "$APP_DIR\logs") {
    Write-Host "   OK -- $APP_DIR\logs exists" -ForegroundColor Green
    Get-ChildItem "$APP_DIR\logs" | ForEach-Object { Write-Host "      $_" }
} else {
    Write-Host "   MISSING -- creating now" -ForegroundColor Red
    New-Item -ItemType Directory -Force -Path "$APP_DIR\logs" | Out-Null
    Write-Host "   Created $APP_DIR\logs" -ForegroundColor Green
}

# -- 2. Check Python ----------------------------------------------------------
Write-Host ""
Write-Host "[2] Python" -ForegroundColor Yellow
$pyPath = $null
try {
    $pyPath = (Get-Command python -ErrorAction Stop).Source
    $pyVer  = python --version 2>&1
    Write-Host "   OK -- $pyPath ($pyVer)" -ForegroundColor Green
} catch {
    Write-Host "   MISSING -- python not found in PATH" -ForegroundColor Red
}

# -- 3. Check required Python packages ----------------------------------------
Write-Host ""
Write-Host "[3] Python packages" -ForegroundColor Yellow
$pkgs = @("flask", "waitress", "requests", "pandas", "openpyxl")
foreach ($pkg in $pkgs) {
    $result = python -c "import $pkg; print($pkg.__version__)" 2>&1
    if ($LASTEXITCODE -eq 0) {
        Write-Host "   OK -- $pkg $result" -ForegroundColor Green
    } else {
        Write-Host "   MISSING -- $pkg  (run: pip install $pkg)" -ForegroundColor Red
    }
}

# -- 4. Test app.py imports ----------------------------------------------------
Write-Host ""
Write-Host "[4] Test app.py import (syntax + import check)" -ForegroundColor Yellow
Set-Location $APP_DIR
$testResult = python -c "import app" 2>&1
if ($LASTEXITCODE -eq 0) {
    Write-Host "   OK -- app.py imports cleanly" -ForegroundColor Green
} else {
    Write-Host "   ERROR -- app.py has import errors:" -ForegroundColor Red
    $testResult | ForEach-Object { Write-Host "   $_" -ForegroundColor Red }
}

# -- 5. Check config.txt -------------------------------------------------------
Write-Host ""
Write-Host "[5] config.txt" -ForegroundColor Yellow
if (Test-Path "$APP_DIR\config.txt") {
    $tokenLine  = Select-String -Path "$APP_DIR\config.txt" -Pattern "META_ACCESS_TOKEN=(.+)"
    $phoneLine  = Select-String -Path "$APP_DIR\config.txt" -Pattern "META_PHONE_NUMBER_ID=(.+)"
    $token   = if ($tokenLine)  { $tokenLine.Matches[0].Groups[1].Value.Trim()  } else { "" }
    $phoneId = if ($phoneLine)  { $phoneLine.Matches[0].Groups[1].Value.Trim()  } else { "" }

    if ($token -and $token -ne "YOUR_TOKEN_HERE") {
        $preview = $token.Substring(0, [Math]::Min(12, $token.Length))
        Write-Host "   OK -- META_ACCESS_TOKEN is set ($preview...)" -ForegroundColor Green
    } else {
        Write-Host "   WARNING -- META_ACCESS_TOKEN not set in config.txt" -ForegroundColor Yellow
    }
    if ($phoneId) {
        Write-Host "   OK -- META_PHONE_NUMBER_ID = $phoneId" -ForegroundColor Green
    } else {
        Write-Host "   WARNING -- META_PHONE_NUMBER_ID not set in config.txt" -ForegroundColor Yellow
    }
} else {
    Write-Host "   MISSING -- config.txt not found" -ForegroundColor Red
}

# -- 6. Check ngrok config -----------------------------------------------------
Write-Host ""
Write-Host "[6] ngrok_do_bot.yml" -ForegroundColor Yellow
if (Test-Path "$APP_DIR\ngrok_do_bot.yml") {
    Write-Host "   OK -- file exists" -ForegroundColor Green
    $hasToken = Select-String -Path "$APP_DIR\ngrok_do_bot.yml" -Pattern "authtoken:\s*\S+" -Quiet
    if ($hasToken) {
        Write-Host "   OK -- authtoken is present" -ForegroundColor Green
    } else {
        Write-Host "   WARNING -- authtoken not found in yml" -ForegroundColor Yellow
    }
} else {
    Write-Host "   MISSING -- copy ngrok_do_bot.yml.example to ngrok_do_bot.yml and fill in your token" -ForegroundColor Red
}

# -- 7. NSSM service registration ----------------------------------------------
Write-Host ""
Write-Host "[7] NSSM service registration" -ForegroundColor Yellow
if (Test-Path $NSSM) {
    foreach ($svc in @("do_bot", "ngrok_do_bot")) {
        $status = & $NSSM status $svc 2>&1
        Write-Host "   $svc : $status"
    }
} else {
    Write-Host "   nssm.exe not found -- run setup_windows.ps1 first" -ForegroundColor Red
}

# -- 8. Show recent error logs -------------------------------------------------
Write-Host ""
Write-Host "[8] Recent error log output" -ForegroundColor Yellow
foreach ($log in @("do_bot_error.log", "ngrok_error.log")) {
    $path = "$APP_DIR\logs\$log"
    if (Test-Path $path) {
        $lines = Get-Content $path -Tail 10 -ErrorAction SilentlyContinue
        if ($lines) {
            Write-Host "   --- $log (last 10 lines) ---" -ForegroundColor Gray
            $lines | ForEach-Object { Write-Host "   $_" }
        } else {
            Write-Host "   $log : empty" -ForegroundColor Gray
        }
    } else {
        Write-Host "   $log : not yet created" -ForegroundColor Gray
    }
}

Write-Host ""
Write-Host "=== Diagnostics complete ===" -ForegroundColor Cyan
Write-Host ""
Write-Host "If do_bot still fails after fixing issues above, run setup_windows.ps1 again" -ForegroundColor Yellow
Write-Host "then: Start-Service do_bot" -ForegroundColor Yellow
