# Network Adapter Manager (WPF)

A Windows-only PowerShell WPF app to **view, filter, and bulk enable/disable network adapters**. It supports quick text filtering, theme switching, selection helpers, **profile save/load** (so your selections persist), and exporting **one-click action scripts** (Enable/Disable). Exports can optionally create a **Scheduled Task + Desktop Shortcut** so future runs don’t trigger UAC prompts.

> **Core script:** `network-adapters-handler.ps1`  
> **Profile file (auto-created):** `adapter-profile.json` (stores which adapters you selected)

---

## Table of Contents

- [Features](#features)
- [User Interface](#ui)
- [Installation](#installation)
- [Usage](#usage)
  - [Filtering & selection](#filtering--selection)
  - [Enable/Disable adapters](#enabledisable-adapters)
  - [Profiles (save/load)](#profiles-saveload)
  - [Export action scripts](#export-action-scripts)
  - [“No-UAC” desktop shortcut](#no-uac-desktop-shortcut)
- [Permissions & compatibility](#permissions--compatibility)
- [Troubleshooting](#troubleshooting)
- [Contributing](#contributing)
- [License](#license)

---

## Features

- **WPF GUI** with **Dark (High Contrast)** and **Light** themes
- **Live filter** box (search Alias/Description)
- **Selection helpers:** Check All (visible), Uncheck All, Invert
- **Bulk actions:** Enable or Disable all checked adapters
- **Profile save/load:** persist selected adapters to `adapter-profile.json`
- **Export scripts:** Generate a portable **Enable** or **Disable** script that
  - resolves adapters robustly (by GUID, Alias, MAC, Description, PnP ID)
  - shows clear success/errors via MessageBoxes
- **Optional “No-UAC” launcher:** create a **Scheduled Task (Highest, Interactive)** and a **Desktop shortcut** that runs the exported script **without** a UAC prompt
- **Self-elevation:** the GUI relaunches as admin in **STA** (required for WPF)

## UI

<img width="1677" height="782" alt="NEtwork Adapter UI" src="https://github.com/user-attachments/assets/1a9e009a-3b01-4cb7-b2ac-9e15a92da7f0" />

## Installation

Quick start is below. For a step-by-step guide and Windows prerequisites, see **[INSTALL.md](INSTALL.md)**.

```powershell
# From a PowerShell prompt on Windows
git clone https://github.com/<you>/network-adapter-manager.git
cd network-adapter-manager

# Run the app (STA is required for WPF)
powershell -NoProfile -ExecutionPolicy Bypass -STA -File .\network-adapters-handler.ps1
```

---

## Usage

### Filtering & selection
- **Filter box** matches **Adapter Name (Alias)** and **Interface Description**.
- Use **Check All (Visible)**, **Uncheck All (Visible)**, or **Invert (Visible)** to work with just the currently filtered rows.

### Enable/Disable adapters
1. Filter (optional), then **check** the adapters you want.
2. Click **Enable Selected** or **Disable Selected**.
3. The app preserves your selections after actions and updates totals/status.

### Profiles (save/load)
- **Save Profile** writes your current selection to `adapter-profile.json`.
- **Load Profile** re-applies those selections.
- If present, a profile is **auto-loaded** at startup.

### Export action scripts
- **Export Enable Script** or **Export Disable Script** to save a standalone `.ps1`.
- The exported script:
  - asks for confirmation,
  - resolves adapters using multiple identifiers (GUID, Alias, MAC, Description, PnP ID),
  - runs the action and shows a result summary.

### “No-UAC” desktop shortcut
- Check **Create “No-UAC” Desktop Shortcut after export** before exporting.
- The app will create:
  - A **Scheduled Task** (Run level: Highest, LogonType: Interactive) that calls PowerShell with `-NoProfile -ExecutionPolicy Bypass -File "<exported.ps1>"`.
  - A **Desktop shortcut** that launches that task—so you can later run the action without a UAC prompt.

---

## Permissions & compatibility
- **Windows only** (uses WPF/PresentationFramework, `Get-NetAdapter`, `Get-CimInstance`, and the ScheduledTasks module).
- **Administrator rights required** to enable/disable adapters and to register the Scheduled Task.
- Tested with **Windows PowerShell 5.1+** on Windows 10/11.

---

## Troubleshooting
- **WPF errors / thread apartment** → make sure you run with `-STA`.
- **Adapters don’t change** → ensure you’re elevated (admin).
- **Scheduled Task creation fails** → verify Task Scheduler is enabled and you have rights; try again from an admin session.
- **No adapters listed** → `Get-NetAdapter` or `Get-CimInstance` may be restricted by policy; check module availability.

---

## Contributing
Contributions are welcome! See **[CONTRIBUTING.md](CONTRIBUTING.md)** for style, testing, and PR guidance.

---

## License
See **LICENSE** for details.
