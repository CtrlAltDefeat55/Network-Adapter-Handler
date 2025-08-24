# Security Policy

## Reporting

If you discover a vulnerability or privacy issue, please open a GitHub issue with **SECURITY** in the title. Avoid publishing exploit details; a maintainer will coordinate next steps and timelines.

## Data & privacy

- The app stores **selection profiles only** in `adapter-profile.json` (next to the script). No telemetry or adapter scan history is persisted.
- Optional “No-UAC” export creates a **Scheduled Task** and a **Desktop shortcut**—both on your local machine only.

## Permissions & risks

- Enabling/disabling network adapters and registering scheduled tasks require **administrator** privileges.
- Exported scripts may be executed without a UAC prompt if you opt in to the **No-UAC** shortcut. Use this convenience responsibly and only on trusted systems.
- Review exported scripts before sharing or running them on other machines.

## Supported versions

The script targets **Windows 10/11 with Windows PowerShell 5.1+**. Newer PowerShell versions on Windows are typically fine as long as WPF assemblies are available.
