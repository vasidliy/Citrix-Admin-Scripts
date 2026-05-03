# Citrix Workspace Hotkey Disabler

[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)

## Overview

This repository provides a Windows Registry File (`.reg`) designed to resolve **hotkey conflicts** between Citrix Workspace (formerly Citrix Receiver) and the published application running inside the Citrix session.

By default, Citrix Workspace reserves a set of keyboard shortcuts (e.g., `SHIFT+F2`, `CTRL+F1`, `ALT+SHIFT+...`) for controlling the session window itself. When a user is working inside a published application that **also uses the same key combination** for its own functions, Citrix intercepts the keypress and prevents it from reaching the remote application.

This configuration performs two essential actions:

1.  **Disables all default Citrix session hotkeys** – This ensures that when a user presses a keyboard shortcut that is normally reserved by Citrix, it is **not intercepted** by the client and is instead passed directly to the published application.
2.  **Sets `TransparentKeyPassthrough` to `FullScreenOnly`** – This fine-tunes the behavior of system key combinations (like `Alt+Tab`) to prevent them from being captured by Citrix when the session is in full-screen mode, allowing the user to switch to local applications without using the mouse.

**Important:** The primary goal of this registry modification is **not** to improve user experience subjectively, but to **enable the correct functionality of published applications** that rely on keyboard shortcuts which overlap with Citrix defaults.

The settings are applied via **Local Machine Policy** (`HKLM\SOFTWARE\Policies\Citrix\...`), meaning they are enforced for **all users** of the workstation and cannot be altered through user preferences.

## Problem Statement

Many line-of-business applications (e.g., SAP GUI, custom terminal emulators, financial trading platforms) use keyboard shortcuts such as:

| Category | Examples | Typical Conflicting Citrix Default |
| :--- | :--- | :--- |
| **Engineering & CAD** | Mathcad Prime, AutoCAD, CATIA, Siemens NX | `Shift+F2` (window toggle), `Ctrl+F1` (Ctrl+Alt+Del) |
| **ERP & Accounting** | SAP GUI, 1C:Enterprise, Microsoft Dynamics | `Shift+F2`, `Ctrl+F1`, `Ctrl+F2` |
| **IDEs & Code Editors** | IntelliJ IDEA, Visual Studio, Eclipse, PyCharm | `Ctrl+F1` (show error description), `Ctrl+F2` (stop) |
| **Graphics & Design** | Adobe Photoshop, Illustrator, CorelDRAW | Various `Ctrl`/`Shift` + Function key combinations |
| **Custom LOB Apps** | In-house developed terminal emulators, reporting tools | Any combination reserved by Citrix (see list below) |


Citrix Workspace, out of the box, intercepts `SHIFT+F2` to toggle full-screen mode. When a user inside the published application presses `SHIFT+F2`, the Citrix client consumes the keystroke, the session toggles full-screen, and **the application never receives the command**. This breaks critical workflows.

This registry file disables the Citrix client's hotkey handler entirely, restoring the keystroke flow to the remote application.

## Detailed Registry Breakdown

### 1. Disabling Citrix Hotkey Interception
**Path:** `HKLM\SOFTWARE\Policies\Citrix\ICA Client\Engine\Lockdown Profiles\All Regions\Lockdown\Client Engine\Hot Keys`

Each Citrix hotkey is defined by a **Shift state** (which modifier keys are pressed) and a **Character** (the main key). By setting both components to `(none)`, the client no longer has a valid trigger for that action.

| Value Name | Data | Result |
| :--- | :--- | :--- |
| `HotKey1Shift` … `HotKey14Shift` | `(none)` | No modifier combination will activate this hotkey. |
| `HotKey1Char` … `HotKey14Char` | `(none)` | No character key will activate this hotkey. |

**Effect:** When a user presses `SHIFT+F2` (or any other default Citrix hotkey), the Citrix Workspace app ignores it completely. The keypress is transmitted via the ICA virtual channel to the published application, where it can be processed as intended by the application vendor.

### 2. Keyboard Passthrough Behavior
**Path:** `HKLM\SOFTWARE\Policies\Citrix\ICA Client\Engine\Lockdown Profiles\All Regions\Lockdown\Virtual Channels\Keyboard`

| Value Name | Data | Default Behavior | Configured Behavior |
| :--- | :--- | :--- | :--- |
| `TransparentKeyPassthrough` | `FullScreenOnly` | Usually `Remote` | **Windowed Mode:** System shortcuts (e.g., `Alt+Tab`) are processed by the remote session.<br>**Full‑Screen Mode:** System shortcuts are processed **locally**, allowing quick exit from the session. |

**Why `FullScreenOnly`?**  
While the main objective is to allow the remote app to receive all keystrokes, we also want to avoid trapping the user inside a full‑screen session. With `FullScreenOnly`, `Alt+Tab` works like a native Windows shortcut when the session is full‑screen, but when the session is windowed, `Alt+Tab` continues to work inside the remote desktop environment.

## Compatibility

- **Client OS:** Windows 10, Windows 11 (x64/x86)
- **Citrix Workspace App:** 1912 LTSR, 2203 LTSR, Current Release (CR)
- **Legacy:** Also compatible with Citrix Receiver 4.x (though the `Lockdown Profiles` path may not be present in very old versions, applying the `.reg` file is harmless).

> **Note:** This registry file writes to the `Policies` hive. It will **override** any user‑level settings in `HKCU` and will take precedence over conflicting settings in Group Policy Objects unless those GPOs explicitly set the same registry keys.

## Installation & Usage

### Method 1: Manual Import (Single Machine)
1.  Download the file `Disable-Citrix-Default-Hotkeys.reg`.
2.  Double‑click the file.
3.  Approve the User Account Control (UAC) prompt.
4.  Confirm the addition to the registry.
5.  **Restart Citrix Workspace** (right‑click the system tray icon → **Exit**, then relaunch).

### Method 2: Deployment via Group Policy (Mass Deployment)
1.  Place the `.reg` file in a network share accessible by target computers.
2.  Create or edit a Group Policy Object.
3.  Navigate to: `Computer Configuration` > `Preferences` > `Windows Settings` > `Registry`.
4.  Create a new **Registry Wizard** item.
5.  Select **Local Computer** and browse to the `.reg` file.
6.  Link the GPO to the desired Organizational Unit.

### Method 3: Command Line (Silent)
```powershell
# PowerShell (Run as Administrator)
Start-Process reg.exe -ArgumentList "import `"C:\Path\To\Disable-Citrix-Default-Hotkeys.reg`"" -Wait -NoNewWindow
```

## Verification

After restarting Citrix Workspace, confirm the settings are active:

1.  Open `regedit.exe`.
2.  Navigate to `HKLM\SOFTWARE\Policies\Citrix\ICA Client\Engine\Lockdown Profiles\All Regions\Lockdown\Client Engine\Hot Keys`.
3.  Verify that the values display as `(none)`.
4.  Navigate to `...\Virtual Channels\Keyboard`.
5.  Verify `TransparentKeyPassthrough` is `FullScreenOnly`.

**Functional Test:** Open a published application that uses a conflicting hotkey (e.g., `SHIFT+F2`). Press the key combination; the application should respond with its native function, and the Citrix session should **not** toggle full‑screen.

## Rollback / Uninstallation

To restore the original Citrix behavior, create a text file named `Restore-Citrix-Hotkeys.reg` with the following content and import it:

```reg
Windows Registry Editor Version 5.00

[-HKEY_LOCAL_MACHINE\SOFTWARE\Policies\Citrix\ICA Client\Engine\Lockdown Profiles\All Regions\Lockdown\Client Engine\Hot Keys]

[HKEY_LOCAL_MACHINE\SOFTWARE\Policies\Citrix\ICA Client\Engine\Lockdown Profiles\All Regions\Lockdown\Virtual Channels\Keyboard]
"TransparentKeyPassthrough"=-
```
*This will delete the custom hotkey policies and reset the keyboard passthrough to the Citrix default.*

## License

MIT License. See `LICENSE` file for details.

## Disclaimer

Modifying the Windows Registry carries inherent risk. While this configuration has been tested in production environments, you should apply it only after appropriate testing in a non‑production environment and ensure a backup or snapshot is available.