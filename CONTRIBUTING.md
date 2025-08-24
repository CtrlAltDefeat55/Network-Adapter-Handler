# Contributing

Thanks for helping improve **Network Adapter Manager (WPF)**! PRs and issues are welcome.

## Project specifics

- **Language/Platform:** PowerShell (Windows), **WPF** UI via `PresentationFramework`
- **Core modules:** `NetAdapter`, `CimCmdlets`, `ScheduledTasks`
- **Admin/STA:** The GUI relaunches elevated and runs **STA** for WPF
- **Selections profile:** Stored next to the script as `adapter-profile.json`

## Dev setup

```powershell
git clone https://github.com/<you>/network-adapter-manager.git
cd network-adapter-manager

# Run with STA & bypass policy for local dev
powershell -NoProfile -ExecutionPolicy Bypass -STA -File .\network-adapters-handler.ps1
```

## Style & quality

- Follow **PowerShell best practices** and **PSScriptAnalyzer** recommendations.
- Keep functions small; prefer clear names and comment-based help where helpful.
- Keep UI text concise and actionable.

### Linting

```powershell
Install-Module PSScriptAnalyzer -Scope CurrentUser
Invoke-ScriptAnalyzer -Path .\network-adapters-handler.ps1 -Recurse
```

## Testing checklist

Before submitting a PR, please verify:

- App launches, shows adapters, and **filter** works
- **Enable/Disable** affects selected adapters and shows status
- **Profile** (save/load) works; auto-load on startup when file exists
- **Export** Enable/Disable scripts run independently and summarize results
- **No-UAC** export option creates a working Scheduled Task + Desktop shortcut

## Commit messages

Use clear, descriptive prefixes when possible:
- `ui:` changes to XAML/controls/themes
- `action:` enable/disable logic
- `export:` generated script / task/shortcut behavior
- `fix:` bug fix or regression
- `docs:` README/INSTALL/etc updates
