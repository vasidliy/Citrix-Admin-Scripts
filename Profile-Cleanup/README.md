# Remove-OldProfiles.ps1

A powerful PowerShell script to automatically clean up old, corrupted, or orphaned user profiles (Citrix UPM, Microsoft Roaming, and reset profiles) from network shares. It supports test mode, quarantine, Active Directory integration, folder redirection handling, detailed logging, and email reports.

## Features

- **Profile type detection** – correctly identifies Citrix UPM, Microsoft Roaming, and reset (`upm_*`) profiles using configurable indicators.
- **Health check** – validates `NTUSER.DAT` existence and size (configurable threshold).
- **Age-based cleanup** – deletes/moves profiles that have not been used for a configurable number of days.
- **New user folder skipping** – optionally skip folders created more recently than the inactivity threshold, unless aggressive actions (delete/quarantine) are configured for corrupted profiles.
- **Active Directory integration** – checks if a user account exists and is enabled; can process per‑source AD domains.
- **Quarantine** – moves profiles to a separate folder instead of deleting them immediately; old quarantine items are automatically removed after a set period.
- **Empty folder cleanup** – removes empty user folders after profile deletion (with exclusion patterns).
- **Folder Redirection (FR) support** – optionally processes redirected folders (Documents, Desktop, etc.) associated with a user when their profile is removed.
- **Orphaned FR handling** – detects and optionally cleans up redirected folders left behind when a profile no longer exists, with configurable modes (report only, delete/move).
- **Intelligent age fallback** – when `NTUSER.DAT` is missing or damaged, can use alternative sources (profile folder or `UPMSettings.ini`) to determine age before deciding action.
- **Granular control over corrupted profiles** – separate actions for missing vs. too small `NTUSER.DAT`.
- **Adaptive profile scanning** – recursively scans subfolders up to a configurable depth using customizable indicators. Works with any profile storage structure (Citrix UPM with platform subfolders, Microsoft Roaming, etc.).
- **Test mode** – dry‑run mode that only logs what would be done, without making any changes.
- **Email reporting** – sends a summary report (plain text in body + HTML attachment).
- **Detailed logging** – writes to a rotating log file with configurable retention.
- **Multi‑source support** – can process multiple profile shares with different settings.
- **Long path support** – handles paths longer than 260 characters and special characters (e.g., `%3A`) using `\\?\UNC\` prefix.

## Requirements

- PowerShell 5.1 or later
- Access to the profile network shares (UNC paths)
- (Optional) ActiveDirectory PowerShell module if AD checks are enabled
- SMTP server for email reports (optional)

## Installation

1. Download the files:
   - `Remove-OldProfiles.ps1`
   - `Remove-OldProfiles.ps1.example.json`
2. Place them in the same directory.
3. Copy `Remove-OldProfiles.ps1.example.json` to `Remove-OldProfiles.ps1.json`.
4. Edit `Remove-OldProfiles.ps1.json` to match your environment.
5. Run the script.

## Understanding Profile Storage Structures

The script expects a layout where **user folders** reside directly under `ProfileRoot`. The name of each folder is treated as the username (after normalisation – removing version suffixes, domain prefixes, etc.). Inside each user folder, the script searches for **profile folders** using configurable indicators.

### Typical Structures

#### Scenario 1: Profiles directly in user folder (Microsoft Roaming)

```
\\fs01\Profiles               <-- ProfileRoot
├── John                      <-- User folder
│   ├── NTUSER.DAT
│   ├── AppData
│   └── Desktop
├── Petrov
│   ├── NTUSER.DAT
│   └── ...
```

#### Scenario 2: Citrix UPM with platform subfolders

```
\\fs01\Profiles               <-- ProfileRoot
├── John                      <-- User folder
│   ├── Win2016v6             <-- Profile folder (Citrix UPM)
│   │   ├── UPM_Profile
│   │   │   └── NTUSER.DAT
│   │   ├── UPMSettings.ini
│   │   └── Pending
│   └── Win10x64              <-- Another profile folder for the same user
│       ├── UPM_Profile
│       │   └── NTUSER.DAT
│       └── UPMSettings.ini
├── Petrov
│   └── Win2019v6
│       ├── UPM_Profile
│       │   └── NTUSER.DAT
│       └── UPMSettings.ini
```

#### Scenario 3: Mixed or custom structures

The adaptive scanner can handle any depth of nesting (up to `MaxProfileScanDepth`) and any set of indicators (files/folders that signal a profile). For example, if your profiles are stored as `\\server\share\username\Profiles\V2\NTUSER.DAT`, you can set `MaxProfileScanDepth = 3` and add `"Profiles\V2\NTUSER.DAT"` to `ProfileIndicators`.

## Configuration Reference

### General Settings

| Field | Type | Default | Description |
|-------|------|---------|-------------|
| `TestMode` | `bool` | `false` | If `true`, no actual file operations are performed. |
| `LogDirectory` | `string` | `"Logs"` | Relative or absolute path for log files. |
| `LogRetentionDays` | `int` | `30` | How many days to keep log files. |
| `ConsoleOutput` | `bool` | `true` | Write log messages to the console. |
| `DetailedLogging` | `bool` | `true` | Include `DEBUG` level messages in the log. |

### Active Directory (Global)

| Field | Type | Default | Description |
|-------|------|---------|-------------|
| `Enabled` | `bool` | `true` | Enable/disable all AD checks. |
| `FQDNDomain` | `string` | – | Fully qualified domain name (e.g., `contoso.com`). |
| `NetBIOSDomain` | `string` | – | NetBIOS domain name (e.g., `CONTOSO`). |
| `ExcludeUsers` | `string[]` | `[]` | Usernames that should **never** be cleaned up. |

### Mail Settings

| Field | Type | Default | Description |
|-------|------|---------|-------------|
| `Enabled` | `bool` | `false` | Send email reports after execution. |
| `SmtpServer` | `string` | – | SMTP server hostname or IP. |
| `Port` | `int` | `25` | SMTP port. |
| `UseSSL` | `bool` | `false` | Use SSL/TLS for SMTP connection. |
| `Credentials.UserName` | `string` | `""` | Optional username for SMTP auth. |
| `Credentials.Password` | `string` | `""` | Optional password for SMTP auth. |
| `From` | `string` | – | Sender email address. |
| `To` | `string[]` | – | Recipient email addresses. |
| `Subject` | `string` | `"[Report] Profile Cleanup"` | Email subject line. |

### Profile Sources (Per Source)

| Field | Type | Default | Description |
|-------|------|---------|-------------|
| `Name` | `string` | – | Descriptive name for this source. |
| `Enabled` | `bool` | `true` | If `false`, this source is skipped. |
| `ProfileRoot` | `string` | – | UNC path to the root folder containing **user folders**. |
| `PatternProfile` | `string` | `"*"` | Wildcard pattern to match user folders. |
| `DaysToDelete` | `int` | `365` | Inactivity threshold in days. |
| `SkipNewUserFolders` | `bool` | `false` | If `true`, skip folders whose creation timestamp is newer than `DaysToDelete` days, **unless** a corrupted profile action (`ActionOnMissingNTUserDat` / `ActionOnTooSmallNTUserDat`) is set to `Delete` or `Quarantine`. |
| `EnableQuarantine` | `bool` | `true` | Move profiles to `QuarantinePath` instead of deleting. |
| `QuarantinePath` | `string` | – | UNC path where quarantined items will be stored. |
| `QuarantineDays` | `int` | `14` | After this many days, items in quarantine are permanently deleted. |
| `ExcludeFolders` | `string[]` | `[]` | Folder name patterns to ignore (e.g., `"Public"`, `"Default"`). |
| `EnableEmptyFolderCleanup` | `bool` | `true` | Delete empty user folders after processing. |
| `EmptyFolderExcludePatterns` | `string[]` | (uses `ExcludeFolders`) | Patterns to exclude from empty folder cleanup. |
| `ActiveDirectory` | `object` | (global) | Per‑source AD override. |
| `FolderRedirectionPaths` | `string[]` | `[]` | Root paths for redirected folders (FR). Subfolders with usernames will be processed together with profiles. |
| `FolderRedirectionExcludePatterns` | `string[]` | `[]` | Patterns for subfolder names inside FR paths to never touch. |
| `ProcessOrphanedFR` | `string`/`bool` | `"Disabled"` | `"Disabled"`, `"ReportOnly"`, or `"Delete"`. |
| `ActionOnMissingNTUserDat` | `string` | `"Quarantine"` | `"Quarantine"`, `"Delete"`, or `"Ignore"`. |
| `ActionOnTooSmallNTUserDat` | `string` | `"Quarantine"` | Same as above. |
| `UseFallbackAgeWhenNTUserMissing` | `bool` | `true` | Use alternative source to determine age when NTUSER.DAT is missing/too small. |
| `FallbackAgeSource` | `string` | `"Auto"` | `"Auto"`, `"Folder"`, or `"UPMSettings"`. |
| `MaxProfileScanDepth` | `int` | `2` | Maximum depth to scan inside a user folder for profile indicators. |
| `ProfileIndicators` | `string[]` | `["UPM_Profile\NTUSER.DAT", "NTUSER.DAT", "UPMSettings.ini"]` | Relative paths or filenames that signal a profile folder. |

#### `SkipNewUserFolders` behaviour in detail

- **Default** – `false`: all user folders are processed regardless of their creation date.
- When `true`:
  - The script checks the **creation time** of the user folder (the top‑level folder under `ProfileRoot`).
  - If the folder was created **less than** `DaysToDelete` days ago, it is normally skipped (a debug message is logged).
  - **Exception**: If the source is configured with `ActionOnMissingNTUserDat` or `ActionOnTooSmallNTUserDat` set to `"Delete"` or `"Quarantine"`, the folder will **still be processed** – even if it is new – because an aggressive action suggests that even recent corrupted profiles must be handled. In this case a debug message indicates that the folder is processed despite `SkipNewUserFolders`.

This parameter is useful to avoid cleaning brand‑new folders that might not yet have a populated profile, while still allowing immediate cleanup of clearly broken profiles (missing/tiny `NTUSER.DAT`) when needed.

## Detailed Explanation of Adaptive Scanning Parameters

### `MaxProfileScanDepth`
Controls how many levels of subfolders the script will descend into when searching for profiles inside a user's folder.
- **0**: only the user folder itself is checked.
- **1**: user folder and its immediate subfolders.
- **2** (default): covers standard Citrix UPM with one platform subfolder (`username\Win2016v6\`).
- Higher values can be set for deeper custom structures.

### `ProfileIndicators`
A list of relative paths (from the scanned folder) that indicate the presence of a profile. When any of these is found, the folder is considered a profile and scanning **stops** for that branch (preventing duplicate detection of subfolders like `UPM_Profile` inside an already identified UPM profile).

**Default indicators**:
- `UPM_Profile\NTUSER.DAT` – Citrix UPM profile.
- `NTUSER.DAT` – Microsoft Roaming or local profile.
- `UPMSettings.ini` – Citrix UPM configuration file (often at the same level as the platform folder).

**Customisation example**: If your profiles are stored as `username\Profile\NTUSER.DAT`, add `"Profile\NTUSER.DAT"` to this list.

## Example Configuration Scenarios

### 1. Standard Citrix UPM with platform subfolders

```json
"ProfileSources": [{
    "Name": "Citrix UPM Share",
    "ProfileRoot": "\\\\fs01\\Profiles",
    "PatternProfile": "*",
    "DaysToDelete": 90,
    "SkipNewUserFolders": false,
    "EnableQuarantine": true,
    "QuarantinePath": "\\\\fs01\\Quarantine",
    "QuarantineDays": 30,
    "MaxProfileScanDepth": 2,
    "ProfileIndicators": ["UPM_Profile\\NTUSER.DAT", "NTUSER.DAT", "UPMSettings.ini"]
}]
```

### 2. Microsoft Roaming profiles directly in user folders

```json
"ProfileSources": [{
    "Name": "Roaming Profiles",
    "ProfileRoot": "\\\\fs02\\RoamingProfiles",
    "PatternProfile": "*",
    "DaysToDelete": 60,
    "SkipNewUserFolders": false,
    "EnableQuarantine": false,
    "MaxProfileScanDepth": 0,
    "ProfileIndicators": ["NTUSER.DAT"]
}]
```

### 3. Mixed environment with Folder Redirection and orphan reporting

```json
"ProfileSources": [{
    "Name": "Mixed Environment",
    "ProfileRoot": "\\\\fs03\\Users",
    "PatternProfile": "*",
    "DaysToDelete": 120,
    "SkipNewUserFolders": true,
    "EnableQuarantine": true,
    "QuarantinePath": "\\\\fs03\\Quarantine",
    "QuarantineDays": 14,
    "MaxProfileScanDepth": 3,
    "FolderRedirectionPaths": [
        "\\\\fs03\\Redirected\\Documents",
        "\\\\fs03\\Redirected\\Desktop"
    ],
    "ProcessOrphanedFR": "ReportOnly",
    "ActionOnMissingNTUserDat": "Ignore",
    "ActionOnTooSmallNTUserDat": "Quarantine",
    "UseFallbackAgeWhenNTUserMissing": true,
    "FallbackAgeSource": "Auto"
}]
```

*In this example, `SkipNewUserFolders = true` prevents processing recently created user folders unless `ActionOnTooSmallNTUserDat` is `Quarantine` – in that case even a new folder with an undersized `NTUSER.DAT` would be processed.*

### 4. Deeply nested custom profile structure

Suppose profiles are stored as:
```
\\server\share\username\Profiles\V2\NTUSER.DAT
```
Configuration:
```json
"MaxProfileScanDepth": 3,
"ProfileIndicators": ["Profiles\\V2\\NTUSER.DAT"]
```

## Usage

```powershell
.\Remove-OldProfiles.ps1                       # Use default config file
.\Remove-OldProfiles.ps1 -ConfigFile prod.json # Custom config
```

Always run with `"TestMode": true` first to review planned actions.

## Troubleshooting

| Issue | Solution |
|-------|----------|
| Profiles not detected | Increase `MaxProfileScanDepth` or adjust `ProfileIndicators`. Enable `DetailedLogging` to see scan results. |
| Duplicate profiles detected | Check that indicators are not being matched in subfolders of already identified profiles. The script stops scanning a branch after a profile is found. If duplicates persist, verify `ProfileIndicators` are specific enough. |
| Access denied | Ensure the account has read/write permissions on the share. Long paths are supported via `\\?\UNC\` prefix. |
| Email not sent | Verify SMTP settings and network connectivity. |

## License

MIT License – feel free to modify and distribute.