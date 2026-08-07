# Running the Lorry Assignment web app 24/7 on Windows

The web app (`web_app.py`) serves the portal at
`http://<this-pc-ip>:8000/login` — e.g. **http://10.0.0.229:8000/login** —
so any phone or PC on the same office WiFi can use it.

To keep it running 24/7 (auto-start on boot, auto-restart on crash) we install
it as a Windows **service** using **NSSM** (Non-Sucking Service Manager).

---

## 1. One-time setup

Open **PowerShell as Administrator** in the project folder
(`C:\Users\Administrator\OneDrive\Documents\GitHub\DO_bot`).

> ⚠️ PowerShell (this version) does **not** support `&&` to join commands.
> Run each line separately, or join with a semicolon `;`.

Install the Python dependencies (waitress = the production server):

```powershell
pip install flask pandas openpyxl waitress
```

Download NSSM from https://nssm.cc/download , unzip it, and copy
`nssm.exe` (the 64-bit one from `win64\`) into the project folder — or note
its full path.

---

## Quick way — one script

In an **Administrator** PowerShell, from the project folder:

```powershell
.\install_service.ps1
```

It auto-detects the real Python (skipping the WindowsApps stub), installs the
service, opens the firewall, and starts it. The manual steps below do the same thing.

## 2. Create the service

Find your Python path first:

```powershell
(Get-Command python).Source
```

Say it prints `C:\Python311\python.exe`. Create the service (adjust paths):

```powershell
.\nssm.exe install LorryAssignment "C:\Python311\python.exe" "C:\Users\Administrator\OneDrive\Documents\GitHub\DO_bot\web_app.py"
.\nssm.exe set LorryAssignment AppDirectory "C:\Users\Administrator\OneDrive\Documents\GitHub\DO_bot"
.\nssm.exe set LorryAssignment Start SERVICE_AUTO_START
.\nssm.exe set LorryAssignment AppStdout "C:\Users\Administrator\OneDrive\Documents\GitHub\DO_bot\logs\web.log"
.\nssm.exe set LorryAssignment AppStderr "C:\Users\Administrator\OneDrive\Documents\GitHub\DO_bot\logs\web.log"
```

Start it:

```powershell
.\nssm.exe start LorryAssignment
```

NSSM auto-restarts the app if it ever crashes, and starts it on every boot.

---

## 3. Open the firewall (so phones can reach it)

Once, as Administrator:

```powershell
New-NetFirewallRule -DisplayName "Lorry Assignment 8000" -Direction Inbound -Protocol TCP -LocalPort 8000 -Action Allow
```

---

## 4. Use it

- On this PC: **http://127.0.0.1:8000/login**
- From a phone/PC on the same WiFi: **http://10.0.0.229:8000/login**
  (if the PC's IP changes, run `ipconfig` and use the new IPv4 address)

Log in with a Company Email + Password from `data\credentials.xlsx`.

---

## Managing the service

```powershell
.\nssm.exe restart LorryAssignment   # after editing credentials or code
.\nssm.exe stop LorryAssignment
.\nssm.exe status LorryAssignment
.\nssm.exe remove LorryAssignment confirm   # uninstall the service
```

- **Change staff / passwords:** edit `data\credentials.xlsx` — takes effect on
  the next login, no restart needed.
- **Update the app after a `git pull`:** `.\nssm.exe restart LorryAssignment`.
- **Logs:** `logs\web.log`.
