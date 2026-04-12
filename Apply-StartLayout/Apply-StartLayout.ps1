<#
.SYNOPSIS
    Applies a customized Start menu and taskbar layout for new users.
    Designed to be triggered by GPO logon scripts and to bypass restrictive ExecutionPolicy settings.
#>

param(
    [string]$XmlPath   # Optional path to LayoutModification.xml (overrides environment variable and default)
)

Start-Transcript -Path C:\Temp\ApplyStartLayout.log -Append

# Determine the XML source file (priority: parameter -> environment variable -> script directory) ---
if ($XmlPath) {
    $xmlSource = $XmlPath
    Write-Host "Using XML from parameter: $xmlSource"
}
elseif ($env:XmlPath) {
    $xmlSource = $env:XmlPath
    Write-Host "Using XML from environment variable: $xmlSource"
}
else {
    $scriptDir = $PSScriptRoot
    if (-not $scriptDir) {
        $scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
    }
    $xmlSource = Join-Path $scriptDir "LayoutModification.xml"
    Write-Host "Looking for XML in script folder: $xmlSource"
}

$flagFile = "$env:APPDATA\StartLayoutApplied.flag"
$xmlDest = "$env:LOCALAPPDATA\Microsoft\Windows\Shell\LayoutModification.xml"

# Always create a "Documents" shortcut (some XML layouts reference it) ---
$docShortcut = "$env:APPDATA\Microsoft\Windows\Start Menu\Programs\Documents.lnk"
if (-not (Test-Path $docShortcut)) {
    Write-Host "Creating 'Documents' shortcut..."
    $destDir = Split-Path $docShortcut -Parent
    if (-not (Test-Path $destDir)) {
        New-Item -ItemType Directory -Path $destDir -Force | Out-Null
    }
    $ws = New-Object -ComObject WScript.Shell
    $sc = $ws.CreateShortcut($docShortcut)
    $sc.TargetPath = '::{450D8FBA-AD25-11D0-98A8-0800361B1103}'  # CLSID for Documents folder
    $sc.Save()
    Write-Host "Shortcut created: $docShortcut"
}
else {
    Write-Host "'Documents' shortcut already exists."
}

# Verify that the XML layout file exists ---
if (-not (Test-Path $xmlSource)) {
    Write-Error "LayoutModification.xml not found: $xmlSource"
    Stop-Transcript
    exit 1
}

# Apply the layout only on the first run (flag file prevents re-application) ---
if (-not (Test-Path $flagFile)) {
    Write-Host "First logon detected. Applying Start layout..."

    # Copy the XML file to the user's local profile folder
    $destDir = Split-Path $xmlDest -Parent
    if (-not (Test-Path $destDir)) {
        New-Item -ItemType Directory -Path $destDir -Force | Out-Null
        Write-Host "Created directory: $destDir"
    }
    Copy-Item -Path $xmlSource -Destination $xmlDest -Force
    Write-Host "XML copied to: $xmlDest"

    # Retrieve the current user's SID (reliable method)
    $userSid = [System.Security.Principal.WindowsIdentity]::GetCurrent().User.Value
    Write-Host "User SID: $userSid"

    # Clear the CloudStore registry key to remove cached Start menu data
    $cloudStorePath = "Registry::HKEY_USERS\$userSid\Software\Microsoft\Windows\CurrentVersion\CloudStore\Store"
    if (Test-Path $cloudStorePath) {
        Remove-Item $cloudStorePath -Force -Recurse
        Write-Host "CloudStore cleared."
    }
    else {
        Write-Host "CloudStore not found."
    }

    # Restart Explorer to apply changes immediately
    Write-Host "Restarting explorer.exe..."
    Stop-Process -Name explorer -Force -ErrorAction SilentlyContinue
    Start-Sleep -Seconds 5
    if (-not (Get-Process explorer -ErrorAction SilentlyContinue)) {
        Start-Process explorer
        Write-Host "Explorer started."
    }

    # Give Windows enough time (30 seconds) to process the layout file,
    # then delete it in the background so it does not interfere with subsequent logons.
    Write-Host "Waiting 30 seconds for Windows to apply the layout..."
    Start-Sleep -Seconds 30

    # Remove the XML file if it still exists
    if (Test-Path $xmlDest) {
        Remove-Item -Path $xmlDest -Force -ErrorAction SilentlyContinue
        Write-Host "XML removed from local profile."
    }

    # Create a flag file to skip layout application on future logons
    New-Item -ItemType File -Path $flagFile -Force | Out-Null
    Write-Host "Flag file created: $flagFile"
}
else {
    Write-Host "Flag file exists. Layout application skipped."
    # Clean up any leftover XML from a previous interrupted run
    if (Test-Path $xmlDest) {
        Remove-Item -Path $xmlDest -Force -ErrorAction SilentlyContinue
        Write-Host "Orphaned XML file removed."
    }
}

Stop-Transcript