# Install & Run — Network Adapter Manager (WPF)

> **Compatibility:** **Windows 10/11** only. Requires **Windows PowerShell 5.1+**, WPF (`PresentationFramework`), and built-in modules: **NetAdapter**, **CimCmdlets**, **ScheduledTasks**.

## 1) Get the code

```powershell
git clone https://github.com/<you>/network-adapter-manager.git
cd network-adapter-manager
```

## 2) Run with the right switches

WPF requires a **Single-Threaded Apartment (STA)** runspace, and you may need to bypass execution policy for local scripts:

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -STA -File .\network-adapters-handler.ps1
```

You can also **Unblock** the file once:

```powershell
Unblock-File .\network-adapters-handler.ps1
```

## 3) Elevation (UAC)

The app will auto-relaunch elevated (admin) if needed. Admin rights are required to enable/disable adapters and to register Scheduled Tasks (for the No-UAC shortcut).

## 4) Exporting action scripts

Use **Export Enable Script** / **Export Disable Script** inside the app to generate a portable `.ps1`.  
Optionally check **Create "No-UAC" Desktop Shortcut after export** to:

- Register a **Scheduled Task** (Highest privileges, Interactive)
- Create a matching **Desktop shortcut** that launches the task (no UAC prompt)

## 5) Common issues

- **“STA required” or WPF errors** → ensure you included `-STA`.
- **Execution policy blocks** → use `-ExecutionPolicy Bypass` or adjust policy per your org.
- **Task creation errors** → confirm Task Scheduler service is running; re-run app as admin.
- **Adapters not changing** → verify elevation and that your account has rights to modify network adapters.
