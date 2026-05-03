# Citrix Admin Scripts

A collection of PowerShell scripts, configuration examples, and utilities for managing **Citrix Virtual Apps & Desktops**, **NetScaler/ADC**, and related components (UPM, profiles, etc.). These tools are designed to simplify the daily tasks of Citrix administrators.

## Contents

- [Profile Cleanup](#profile-cleanup) — removal of stale, corrupted, or orphaned user profiles (Citrix UPM, Microsoft Roaming) with quarantine, Active Directory integration, and email reporting.
- [Apply Start Layout](#apply-start-layout) — configuration of the Start menu and taskbar for new Windows 10/11 users.
- [Citrix Workspace Hotkey Disabler](#citrix-workspace-hotkey-disabler) — disabling Citrix Workspace hotkey interception.
- [Usage](#usage)
- [License](#license)

---

## Profile Cleanup

The `Remove-OldProfiles.ps1` script is a powerful tool for automatically cleaning up old, corrupted, or orphaned user profiles (Citrix UPM, Microsoft Roaming, reset profiles) from network shares. It supports a test mode, quarantine, Active Directory integration, folder redirection handling, detailed logging, and email reporting.

[Learn more »](Profile-Cleanup/README.md)

## Apply Start Layout

Contains scripts and a layout template that configure the Start menu and taskbar for **new users** on Windows 10 and Windows 11. The solution is intended for deployment through Group Policy (GPO) as logon scripts or for manual execution in the user context.

[Learn more »](Apply-StartLayout/README.md)

## Citrix Workspace Hotkey Disabler

A Windows registry file (`.reg`) designed to resolve hotkey conflicts between Citrix Workspace and published applications. It disables Citrix hotkey interception and enables transparent key passing to ensure applications that use the same key combinations work correctly.

[Learn more »](Citrix-Workspace-Hotkey-Disabler/README.md)

## Usage

Each subfolder contains its own `README.md` file with detailed setup and usage instructions for the respective tool.

## License

This project is distributed under the MIT License. See the [LICENSE](LICENSE) file for details.