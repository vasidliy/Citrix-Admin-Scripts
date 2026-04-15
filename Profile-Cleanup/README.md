# Remove-OldProfiles.ps1

A powerful PowerShell script to automatically clean up old, corrupted, or orphaned user profiles (Citrix UPM, Microsoft Roaming, and reset profiles) from network shares. It supports test mode, quarantine, Active Directory integration, detailed logging, and email reports.

## Features

- **Profile type detection** – correctly identifies Citrix UPM, Microsoft Roaming, and reset (upm_*) profiles.
- **Health check** – validates `NTUSER.DAT` existence and size (optional).
- **Age-based cleanup** – deletes/moves profiles that have not been used for a configurable number of days.
- **Active Directory integration** – checks if a user account exists and is enabled; can process per‑source AD domains.
- **Quarantine** – moves profiles to a separate folder instead of deleting them immediately; old quarantine items are automatically removed after a set period.
- **Empty folder cleanup** – removes empty user folders after profile deletion (with exclusion patterns).
- **Test mode** – dry‑run mode that only logs what would be done, without making any changes.
- **Email reporting** – sends a summary report (text and optional HTML) with attached log files.
- **Detailed logging** – writes to a rotating log file with configurable retention.
- **Multi‑source support** – can process multiple profile shares with different settings.

## Requirements

- PowerShell 5.1 or later
- Access to the profile network shares (UNC paths)
- (Optional) ActiveDirectory PowerShell module if AD checks are enabled
- SMTP server for email reports (optional)

## Installation

1. Download both files:
   - `Remove-OldProfiles.ps1`
   - `Remove-OldProfiles.ps1.example.json`
2. Place them in the same directory (or adjust the `-ConfigFile` parameter).
3. Copy `Remove-OldProfiles.ps1.example.json` to `Remove-OldProfiles.ps1.json`.
4. Edit the new `Remove-OldProfiles.ps1.json` file to match your environment.
5. Run the script.

## Configuration

The script uses a JSON configuration file. Below is an example with explanations:

```json
{
    "GeneralSettings": {
        "TestMode": false,                    // true = dry run, false = real execution
        "LogDirectory": "Logs",               // subfolder for logs (relative to script)
        "LogRetentionDays": 30,               // delete logs older than this
        "ConsoleOutput": true,                // show output in console
        "DetailedLogging": true               // write DEBUG level messages
    },
    "ActiveDirectory": {
        "Enabled": true,                      // global AD checks
        "FQDNDomain": "domain.company.com",
        "NetBIOSDomain": "DOMAIN",
        "ExcludeUsers": [ "Administrator", "Guest", "krbtgt" ]
    },
    "MailSettings": {
        "Enabled": true,
        "SmtpServer": "mail.company.com",
        "Port": 25,
        "UseSSL": false,
        "Credentials": { "UserName": "", "Password": "" }, // optional
        "From": "notify@company.com",
        "To": [ "support@company.com" ],
        "Subject": "[Report] Profile Cleanup"
    },
    "ProfileSources": [
        {
            "Name": "fserver01",
            "Enabled": true,
            "ProfileRoot": "\\\\fserver01\\Profiles",
            "PatternProfile": "*",            // wildcard for user folders
            "DaysToDelete": 365,              // inactivity threshold (days)
            "EnableQuarantine": false,        // if false, profiles are deleted directly
            "QuarantinePath": "\\\\fserver01\\Profiles\\Pending",
            "QuarantineDays": 14,             // auto‑delete from quarantine after N days
            "ExcludeFolders": [               // folders to ignore (wildcards allowed)
                "Public", "Default", "Temp*", "*.bak", "!ThinAppCap", "~snapshot"
            ],
            "EnableEmptyFolderCleanup": false,
            "EmptyFolderExcludePatterns": [   // optional; uses ExcludeFolders if missing
                "Public", "Default", "Temp*"
            ],
            "ActiveDirectory": {               // per‑source AD settings (override global)
                "Enabled": true,
                "FQDNDomain": "domain.company.com",
                "NetBIOSDomain": "DOMAIN",
                "ExcludeUsers": [ "Administrator", "Guest" ]
            }
        }
    ]
}
```

### Explanation of key fields

| Field | Description |
|-------|-------------|
| `TestMode` | If `true`, no actual move/delete operations are performed. |
| `ActiveDirectory.Enabled` | When enabled, the script checks if the user exists and is enabled. Disabled users or non‑existent users cause profile cleanup. |
| `ProfileSources[].DaysToDelete` | Number of days of inactivity (based on `NTUSER.DAT` last write time) before a profile is considered old. |
| `EnableQuarantine` | If `true`, profiles are moved to `QuarantinePath` instead of being deleted. If `false`, they are deleted immediately. |
| `QuarantineDays` | Profiles older than this in the quarantine folder are permanently deleted. |
| `ExcludeFolders` | Folders (wildcards allowed) that are completely ignored (e.g., `Public`, `Default`). |
| `EnableEmptyFolderCleanup` | If `true`, removes empty user folders after all profiles inside have been processed. |
| `EmptyFolderExcludePatterns` | Patterns to exclude from empty folder cleanup (if not set, falls back to `ExcludeFolders`). |
| `ActiveDirectory` (per source) | Optional per‑source AD settings. If missing, the global `ActiveDirectory` section is used. |

## How It Works

1. **Load configuration** – reads the JSON file.
2. **For each enabled profile source**:
   - Scan `ProfileRoot` for folders matching `PatternProfile`, excluding `ExcludeFolders` and the quarantine folder itself.
   - For each user folder, call `Find-UserProfiles` to detect one or more profiles inside:
     - `CitrixUPM` – contains `UPM_Profile` folder or `UPMSettings.ini`.
     - `MicrosoftRoaming` – contains `NTUSER.DAT` or `AppData`/`Desktop`.
     - `ResetCitrixProfile` – folder name matches `upm_*` but not `UPM_Profile`.
     - `NotDefined` – empty folder or no recognisable indicators.
   - For each profile, determine if it should be cleaned up based on:
     - **Health** – `NTUSER.DAT` missing or too small → `PROFILE_CORRUPTED`
     - **Age** – `NTUSER.DAT` last write time older than `DaysToDelete` → `PROFILE_OLD`
     - **AD status** – user not found or disabled → `USER_NOT_FOUND` / `USER_DISABLED`
     - **Reset profiles** – always moved/deleted (reason `RESET_PROFILE`)
   - If a reason exists, the profile is either **moved to quarantine** (if `EnableQuarantine = true`) or **deleted**.
3. **Cleanup**:
   - Remove old profiles from quarantine older than `QuarantineDays`.
   - Delete empty user folders (if enabled).
4. **Logging & Reporting**:
   - Write detailed logs to `LogDirectory`.
   - Generate a text report and an optional HTML report with a table of actions.
   - Send an email summary if `MailSettings.Enabled = true`.
5. **Log rotation** – deletes log files older than `LogRetentionDays`.

## Usage

### Basic execution (using default config file in script directory)

```powershell
.\Remove-OldProfiles.ps1
```

### Specify a different configuration file

```powershell
.\Remove-OldProfiles.ps1 -ConfigFile "C:\MyConfigs\cleanup_prod.json"
```

### Test mode (dry run)

Set `"TestMode": true` in the JSON file, or temporarily change it. In test mode the script will:

- Log all actions as `[TEST]`
- Show which profiles would be moved/deleted
- **Not** modify any files or folders

Always run in test mode first to verify the configuration.

## Active Directory Integration

The script caches AD queries per user per domain to improve performance. For each profile, it checks:

- Whether the user account exists in the specified domain.
- Whether the account is enabled (if it exists).

If the user does **not** exist or is **disabled**, the profile is cleaned up (provided other conditions like age are not already met).

To disable AD checks for a source, set `"ActiveDirectory.Enabled": false` (or globally).

## Quarantine

Quarantine is a safety mechanism. Instead of deleting a profile immediately, the script moves it to a dedicated folder (e.g., `Pending`). This allows you to recover profiles if needed. The quarantine folder is automatically cleaned of entries older than `QuarantineDays`.

If you set `"EnableQuarantine": false`, profiles are deleted directly – use with caution.

## Email Reports

The script can send an email report after execution. The report includes:

- Summary statistics (number of profiles found, moved, deleted, etc.)
- List of processed sources
- Attachments: text report, HTML report (if any profiles were processed), and the log file

Configure `MailSettings` with your SMTP server and recipients. Credentials are optional.

## Logging

Logs are stored in the `LogDirectory` (relative to the script). Each execution creates a new log file: `ProfileCleanup_YYYYMMDD_HHmmss.log`. Log entries include timestamps, level (INFO/WARN/ERROR/DEBUG), and the message.

Old logs are automatically deleted after `LogRetentionDays`.

## Example Scenarios

### 1. Clean up Citrix UPM profiles older than 90 days, move them to quarantine

```json
"ProfileSources": [{
    "Name": "Citrix UPM Share",
    "Enabled": true,
    "ProfileRoot": "\\\\fs\\Profiles",
    "DaysToDelete": 90,
    "EnableQuarantine": true,
    "QuarantinePath": "\\\\fs\\Quarantine",
    "QuarantineDays": 30
}]
```

### 2. Delete Microsoft Roaming profiles for disabled AD users immediately (no quarantine)

```json
"EnableQuarantine": false,
"ActiveDirectory": { "Enabled": true }
```

### 3. Test mode for a new source

Set `"TestMode": true` in `GeneralSettings` and run the script. Review the logs and the HTML report to see what would happen.

## Troubleshooting

| Issue | Possible solution |
|-------|-------------------|
| Script cannot access network share | Ensure the account running the script has read/write permissions on the UNC path. |
| Active Directory queries fail | Verify that the AD module is installed and the domain is reachable. Use `FQDNDomain` to target a specific domain controller. |
| Profiles are not being detected | Check `PatternProfile` and `ExcludeFolders`. Enable `DetailedLogging` to see why folders are skipped. |
| Email not sent | Confirm SMTP server, port, and firewall settings. Check credentials if required. |
| Quarantine folder not created | The script creates the folder automatically if it does not exist. Ensure the parent path is accessible. |

## Best Practices

- **Always test** with `TestMode = true` before running in production.
- Use a dedicated service account with least privilege (read/write on profile shares, read-only for AD).
- Regularly review logs and quarantine folders.
- Consider scheduling the script via Task Scheduler with appropriate credentials.

## License

This script is provided as‑is under the MIT License. Feel free to modify and distribute.