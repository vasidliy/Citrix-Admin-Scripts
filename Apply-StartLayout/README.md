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

### Understanding the XML structure

The Start menu layout is defined within the `<defaultlayout:StartLayout>` element. Tiles are placed on a **grid** where:

- **`GroupCellWidth`** (on the `<defaultlayout:StartLayout>` element) defines the width of each tile group in **tile units**. The default value is `6`, meaning each group can accommodate up to 3 medium tiles (2×2) per row.
- Each `<start:DesktopApplicationTile>` has the following key attributes:
  - **`Size`**: tile dimensions. Common values:
    - `2x2` – Medium tile (occupies 2 columns × 2 rows).
    - `4x2` – Wide tile.
    - `4x4` – Large tile.
  - **`Column`** and **`Row`**: the **zero‑based** coordinates of the tile's **top‑left corner** within the group.
  - **`DesktopApplicationLinkPath`** or **`DesktopApplicationID`**: specifies which application to launch.

**Example interpretation**  
The snippet below places a medium Outlook tile at the top‑left corner of the group:
```xml
<start:DesktopApplicationTile Size="2x2" Column="0" Row="0" DesktopApplicationLinkPath="%ALLUSERSPROFILE%\Microsoft\Windows\Start Menu\Programs\Outlook.lnk" />
```
- **Column="0" Row="0"** → top‑left corner.
- **Size="2x2"** → tile spans columns 0‑1 and rows 0‑1.

### 📐 Visual grid representation of the provided example

The table below shows how the tiles from the included `LayoutModification.xml` are arranged on a `GroupCellWidth="6"` grid.  
**Column 0, Row 0 is the top‑left cell.** Each tile occupies a 2×2 block.

| Row | Col 0 | Col 1 | Col 2 | Col 3 | Col 4 | Col 5 |
|-----|-------|-------|-------|-------|-------|-------|
| **0** | Outlook (0,0) 2x2 | Outlook (cont.) | Documents (2,0) 2x2 | Documents (cont.) | This PC (4,0) 2x2 | This PC (cont.) |
| **1** | Outlook (cont.) | Outlook (cont.) | Documents (cont.) | Documents (cont.) | This PC (cont.) | This PC (cont.) |
| **2** | Edge (0,2) 2x2 | Edge (cont.) | Firefox (2,2) 2x2 | Firefox (cont.) | IE (4,2) 2x2 | IE (cont.) |
| **3** | Edge (cont.) | Edge (cont.) | Firefox (cont.) | Firefox (cont.) | IE (cont.) | IE (cont.) |
| **4** | Excel (0,4) 2x2 | Excel (cont.) | PowerPoint (2,4) 2x2 | PowerPoint (cont.) | Word (4,4) 2x2 | Word (cont.) |
| **5** | Excel (cont.) | Excel (cont.) | PowerPoint (cont.) | PowerPoint (cont.) | Word (cont.) | Word (cont.) |

> **Note:** `(cont.)` indicates cells covered by the same 2×2 tile.  
> For example, Outlook starts at `Column="0" Row="0"` and also occupies `(1,0)`, `(0,1)`, and `(1,1)`.

### 📌 Taskbar pin list structure

The taskbar shortcuts are defined inside the `<CustomTaskbarLayoutCollection>` section:

```xml
<CustomTaskbarLayoutCollection>
    <defaultlayout:TaskbarLayout>
        <taskbar:TaskbarPinList>
            <taskbar:DesktopApp DesktopApplicationID="Microsoft.Windows.Explorer"/>
            <taskbar:DesktopApp DesktopApplicationLinkPath="%ALLUSERSPROFILE%\Microsoft\Windows\Start Menu\Programs\Outlook.lnk" />
            <taskbar:DesktopApp DesktopApplicationLinkPath="%ALLUSERSPROFILE%\Microsoft\Windows\Start Menu\Programs\Firefox.lnk" />
            <taskbar:DesktopApp DesktopApplicationLinkPath="%ALLUSERSPROFILE%\Microsoft\Windows\Start Menu\Programs\Microsoft Edge.lnk" />
        </taskbar:TaskbarPinList>
    </defaultlayout:TaskbarLayout>
</CustomTaskbarLayoutCollection>
```

**How it works:**
- Each `<taskbar:DesktopApp>` element pins **one application** to the taskbar.
- The order of elements in the XML determines the **left‑to‑right order** on the taskbar.
- You can use either:
  - `DesktopApplicationID` – for built‑in Windows apps (File Explorer, Control Panel, etc.).
  - `DesktopApplicationLinkPath` – for any `.lnk` shortcut file (usually located in `%ALLUSERSPROFILE%\Microsoft\Windows\Start Menu\Programs\`).

**Example taskbar order (as shown above):**
1. File Explorer (`Microsoft.Windows.Explorer`)
2. Outlook
3. Firefox
4. Microsoft Edge

> **Note:** The user can later unpin or rearrange these icons. This layout is applied only **once** (on first logon).

### Important notes for customization
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