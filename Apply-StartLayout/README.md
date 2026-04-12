# Apply Start Layout for Windows 10 / 11

This folder contains scripts and a layout template that customize the Start menu and taskbar for **new users** on Windows 10 and Windows 11.  
The solution is designed to be deployed via Group Policy (GPO) logon scripts or run manually in a user context.

## 📁 Files

| File | Description |
|------|-------------|
| `Apply-StartLayout.bat` | Wrapper batch file that invokes the PowerShell script, bypassing execution policy restrictions. |
| `Apply-StartLayout.ps1` | PowerShell script that applies the layout once per user. It copies the XML file, clears the Start cache, restarts Explorer, and sets a flag to prevent re‑application. |
| `LayoutModification.xml` | The layout definition file (Start tiles and taskbar pins). Customize this file to match your organization’s requirements. |

## 🔧 How It Works

1. The batch file passes the XML path to the PowerShell script and runs it **in memory** to avoid the `AllSigned` execution policy requirement.
2. On first logon for a user:
   - A "Documents" shortcut is created (referenced by the layout).
   - The XML file is copied to `%LOCALAPPDATA%\Microsoft\Windows\Shell\LayoutModification.xml`.
   - The CloudStore registry key is cleared to remove any cached Start layout.
   - Explorer is restarted to apply the new layout.
   - A flag file (`%APPDATA%\StartLayoutApplied.flag`) is created so the layout is not reapplied on subsequent logons.
3. After a 30‑second delay (allowing Windows to fully process the XML), the file is deleted from the profile to avoid interference with future user customizations.

## 🚀 Deployment Options

### Group Policy Logon Script
- Place all three files in a network share accessible by all users.
- In Group Policy Management Editor, navigate to **User Configuration → Windows Settings → Scripts (Logon/Logoff) → Logon**.
- Add `Apply-StartLayout.bat` as a logon script.

### Manual Execution
Run the batch file **as the target user**:
```cmd
Apply-StartLayout.bat
```
The script must be executed with user privileges (no administrator rights required).

## ⚙️ Customizing the Layout

Edit `LayoutModification.xml` to change the pinned tiles or taskbar shortcuts.  
**Important notes:**
- Use **DesktopApplicationLinkPath** with valid `.lnk` shortcuts that exist for all users (e.g., `%ALLUSERSPROFILE%\Microsoft\Windows\Start Menu\Programs\...`).
- The script automatically creates a `Documents.lnk` shortcut in the user’s Start Menu Programs folder. If your layout references it, keep that step enabled.
- Test the layout on a clean user profile before mass deployment.

> **Example configuration notice:**  
> The provided `LayoutModification.xml` is an **example** that pins Outlook, Firefox, Microsoft Edge, and Office 2019 applications.  
> Replace `Firefox.lnk` and `Outlook.lnk` with the actual shortcut names from your environment, or update `DesktopApplicationID` for different Office versions.

## 📋 Requirements

- Windows 10 / Windows 11 (any edition).
- PowerShell execution policy can be restricted – the batch file uses `Invoke-Expression` to run the script without touching the policy.
- The user must have write permissions to their own `%APPDATA%` and `%LOCALAPPDATA%` folders (always true for standard users).

## 📄 Logging

The PowerShell script writes a transcript log to `C:\Temp\ApplyStartLayout.log`.  
Ensure the `C:\Temp` directory exists or modify the path inside the script if necessary.

## 🔍 Troubleshooting

| Symptom | Possible Solution |
|---------|-------------------|
| Layout does not apply | Check the transcript log for errors. Verify that `LayoutModification.xml` is accessible and well‑formed. |
| Start menu tiles appear as blank squares | The shortcuts referenced in the XML do not exist for the user. Ensure the `.lnk` files are present in the All Users Start Menu. |
| Script runs on every logon | The flag file was not created. Check write permissions to `%APPDATA%`. |
| Explorer does not restart | Manually restart Explorer or log off and on again. |

## 📝 Notes

- The layout is applied **only once per user**; after that, users can customize their Start menu freely.
- This method does **not** enforce a mandatory layout – for lockdown scenarios, consider using Group Policy **Start Layout** settings instead.
- The script deletes the XML file from the user’s profile after application. If you need to retain it, remove the deletion block in the PowerShell script.

## 📄 License

This script is provided as‑is under the MIT License.
You are free to modify and distribute it.