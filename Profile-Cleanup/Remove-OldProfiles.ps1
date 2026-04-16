#requires -version 5.1

<#
.SYNOPSIS
    Script to delete old profiles with specified profile type and test mode.
    In test mode (TestMode: true) the script only checks but does not move/delete profiles.
    Supports email reporting, folder exclusions, and robocopy for moving.

.DESCRIPTION
    This script scans user profile folders (Citrix UPM, Microsoft Roaming, reset profiles),
    checks profile health (NTUSER.DAT), verifies user status in Active Directory,
    and moves/deletes old or invalid profiles according to the configuration.
    It supports quarantine, empty folder cleanup, folder redirection handling,
    and detailed logging.

.PARAMETER ConfigFile
    Path to the JSON configuration file. Default: "Remove-OldProfiles.ps1.json" in the script directory.

.EXAMPLE
    .\Remove-OldProfiles.ps1 -ConfigFile "C:\Configs\mySettings.json"
    Runs the script with a custom configuration file.

.NOTES
    Requires: PowerShell 5.1, ActiveDirectory module (optional), network access to profile shares.
#>

param(
    [string]$ConfigFile = "Remove-OldProfiles.ps1.json"
)

# ========== SETTINGS ==========
$ScriptPath = Split-Path -Parent $MyInvocation.MyCommand.Path
if ([string]::IsNullOrEmpty($ScriptPath)) {
    $ScriptPath = $PWD.Path
}

# If the path is not absolute, look in the script folder
if (-not $ConfigFile.Contains('\') -and -not $ConfigFile.Contains('/')) {
    $ConfigFile = Join-Path $ScriptPath $ConfigFile
}

# Check configuration file
if (-not (Test-Path $ConfigFile)) {
    Write-Host "ERROR: Configuration file not found: $ConfigFile" -ForegroundColor Red
    exit 1
}

# Load configuration
try {
    $Config = Get-Content $ConfigFile -Raw -Encoding UTF8 | ConvertFrom-Json
}
catch {
    Write-Host "ERROR: Failed to load configuration: $_" -ForegroundColor Red
    exit 1
}

# Settings (global)
$GeneralSettings = $Config.GeneralSettings
$GlobalADSettings = $Config.ActiveDirectory  # will be used as defaults
$MailSettings = $Config.MailSettings
$userProfileSources = $Config.ProfileSources | Where-Object { $_.Enabled -eq $true }

# Test mode
$TestMode = if ($GeneralSettings.TestMode) { $true } else { $false }

# ========== LOGGING ==========
$LogDir = Join-Path $ScriptPath $GeneralSettings.LogDirectory
$LogFile = Join-Path $LogDir "ProfileCleanup_$(Get-Date -Format 'yyyyMMdd_HHmmss').log"

if (-not (Test-Path $LogDir)) {
    New-Item -Path $LogDir -ItemType Directory -Force | Out-Null
}

function Write-Log {
    param(
        [string]$Message,
        [string]$Level = "INFO"
    )
    
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $testPrefix = if ($TestMode) { "[TEST] " } else { "" }
    $logMessage = "[$timestamp] [$Level] ${testPrefix}$Message`n"
    
    # Atomic write with retries
    $maxRetries = 5
    $retryDelay = 200
    $retryCount = 0
    $written = $false
    
    while (-not $written -and $retryCount -lt $maxRetries) {
        try {
            [System.IO.File]::AppendAllText($LogFile, $logMessage, [System.Text.UTF8Encoding]::new($false))
            $written = $true
        }
        catch {
            $retryCount++
            if ($retryCount -ge $maxRetries) {
                Write-Host "Failed to write to log after $maxRetries attempts: $_" -ForegroundColor Red
            }
            else {
                Start-Sleep -Milliseconds $retryDelay
            }
        }
    }
    
    if ($GeneralSettings.ConsoleOutput) {
        switch ($Level) {
            "ERROR" { Write-Host $logMessage.Trim() -ForegroundColor Red }
            "WARN" { Write-Host $logMessage.Trim() -ForegroundColor Yellow }
            "INFO" { Write-Host $logMessage.Trim() -ForegroundColor White }
            "DEBUG" { 
                if ($GeneralSettings.DetailedLogging -or $TestMode) {
                    Write-Host $logMessage.Trim() -ForegroundColor Gray 
                }
            }
        }
    }
}

# ========== EMAIL FUNCTION ==========
function Send-ReportEmail {
    param(
        [string]$ReportFile,
        [string]$HtmlReportFile,
        [string]$LogFile,
        [hashtable]$Stats,
        [bool]$TestMode
    )
    
    if (-not $MailSettings.Enabled) {
        Write-Log "Email sending is disabled in settings" -Level "DEBUG"
        return
    }
    
    try {
        $required = @('SmtpServer', 'From', 'To')
        foreach ($r in $required) {
            if (-not $MailSettings.$r) {
                throw "Required parameter MailSettings.$r is missing"
            }
        }
        
        $subject = $MailSettings.Subject
        if (-not $subject) {
            $subject = "Profile Cleanup Report - $(Get-Date -Format 'yyyy-MM-dd')"
        }
        if ($TestMode) {
            $subject = "[TEST] " + $subject
        }
        
        $body = @"
Profile Cleanup Report
Execution Date: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')
Execution Mode: $(if($TestMode){'TESTING'}else{'REAL'})

SUMMARY:
- Sources processed: $($Stats.SourcesProcessed)/$($Stats.TotalSources)
- Profiles found: $($Stats.TotalProfilesFound)
  * Citrix UPM: $($Stats.CitrixProfilesFound)
  * Microsoft Roaming: $($Stats.RoamingProfilesFound)
  * Reset Citrix: $($Stats.ResetProfilesFound)
  * Undefined/empty profiles: $($Stats.NotDefinedProfilesFound)

$(
if ($TestMode) {
    "TEST MODE - PROFILES WERE NOT PROCESSED:`n"
    "Profiles to process: $($Stats.ProfilesToMoveCount + $Stats.ProfilesToDeleteCount)"
    "  * To move: $($Stats.ProfilesToMoveCount)"
    "  * To delete: $($Stats.ProfilesToDeleteCount)"
    "  * By age: $($Stats.ProfilesToMoveByAge)"
    "  * By AD status: $($Stats.ProfilesToMoveByAD)"
    "  * By corruption: $($Stats.ProfilesToMoveByCorruption)"
    "  * Reset profiles: $($Stats.ResetProfilesMoved)`n"
    "Folder Redirection to process: $($Stats.FolderRedirectionToMoveCount + $Stats.FolderRedirectionToDeleteCount + $Stats.OrphanedFRReportedCount)"
    "  * To move: $($Stats.FolderRedirectionToMoveCount)"
    "  * To delete: $($Stats.FolderRedirectionToDeleteCount)"
    "  * Orphaned reported: $($Stats.OrphanedFRReportedCount)`n"
    "Empty folders to delete: $($Stats.EmptyFolders)`n"
} else {
    "REAL MODE - RESULTS:`n"
    "Moved profiles: $($Stats.ProfilesMovedSuccess)/$($Stats.ProfilesToMoveCount) (success/total)"
    "Deleted profiles: $($Stats.ProfilesDeletedSuccess)/$($Stats.ProfilesToDeleteCount) (success/total)"
    "  * Quarantine cleanup: $($Stats.OldQuarantineProfiles) profiles removed"
    "  * By age: $($Stats.ProfilesToMoveByAge)"
    "  * By AD status: $($Stats.ProfilesToMoveByAD)"
    "  * By corruption: $($Stats.ProfilesToMoveByCorruption)"
    "  * Reset profiles: $($Stats.ResetProfilesMoved)"
    "  * Move errors: $($Stats.ProfilesMoveFailed)"
    "  * Delete errors: $($Stats.ProfilesDeleteFailed)`n"
    "Folder Redirection:"
    "  * Moved: $($Stats.FolderRedirectionMovedSuccess)/$($Stats.FolderRedirectionToMoveCount) (success/total)"
    "  * Deleted: $($Stats.FolderRedirectionDeletedSuccess)/$($Stats.FolderRedirectionToDeleteCount) (success/total)"
    "  * Orphaned processed: $($Stats.OrphanedFRProcessedCount) (moved/deleted)"
    "  * Orphaned reported only: $($Stats.OrphanedFRReportedCount)"
    "  * Move errors: $($Stats.FolderRedirectionMoveFailed)"
    "  * Delete errors: $($Stats.FolderRedirectionDeleteFailed)`n"
    "Empty folders deleted: $($Stats.EmptyFolders)`n"
}
)

Source details:
$($userProfileSources | Where-Object { $_.Enabled } | ForEach-Object {
    $quarantineStatus = if ($_.PSObject.Properties.Name -contains 'EnableQuarantine' -and $_.EnableQuarantine -eq $false) { 'quarantine disabled' } else { 'quarantine enabled' }
    "- $($_.Name): $($_.ProfileRoot) (pattern: $($_.PatternProfile), $quarantineStatus)"
} | Out-String)

Full report and log are attached.
"@
        
        $attachments = @()
        if (Test-Path $ReportFile) { $attachments += $ReportFile }
        if ($HtmlReportFile -and (Test-Path $HtmlReportFile)) { $attachments += $HtmlReportFile }
        if (Test-Path $LogFile) { $attachments += $LogFile }
        
        $mailParams = @{
            SmtpServer  = $MailSettings.SmtpServer
            Port        = if ($MailSettings.Port) { $MailSettings.Port } else { 25 }
            UseSsl      = if ($MailSettings.UseSSL) { $MailSettings.UseSSL } else { $false }
            From        = $MailSettings.From
            To          = $MailSettings.To
            Subject     = $subject
            Body        = $body
            Attachments = $attachments
            ErrorAction = 'Stop'
            Encoding    = [System.Text.Encoding]::UTF8
        }
        
        if ($MailSettings.Credentials.UserName -and $MailSettings.Credentials.Password) {
            $securePass = ConvertTo-SecureString -String $MailSettings.Credentials.Password -AsPlainText -Force
            $cred = New-Object System.Management.Automation.PSCredential ($MailSettings.Credentials.UserName, $securePass)
            $mailParams.Credential = $cred
        }
        
        Send-MailMessage @mailParams
        Write-Log "Report successfully sent to: $($MailSettings.To -join ', ')" -Level "INFO"
    }
    catch {
        Write-Log "Error sending email: $_" -Level "ERROR"
    }
}

# ========== AD CHECK FUNCTION (with caching, per-source domain support) ==========
$script:adCache = @{}

function Get-UserADStatus {
    param(
        [string]$UserName,
        [PSObject]$ADConfig   # ActiveDirectory settings object for this source
    )
    
    if (-not $ADConfig.Enabled) {
        return @{ Exists = $true; Enabled = $true }
    }
    
    # Excluded users (domain-specific)
    if ($ADConfig.ExcludeUsers -and ($UserName -in $ADConfig.ExcludeUsers)) {
        return @{ Exists = $true; Enabled = $true }
    }
    
    # Cache key: domain + user (domain can be FQDN or NetBIOS)
    $domainKey = if ($ADConfig.FQDNDomain) { $ADConfig.FQDNDomain } else { $ADConfig.NetBIOSDomain }
    $cacheKey = "$domainKey\$UserName"
    
    if ($script:adCache.ContainsKey($cacheKey)) {
        return $script:adCache[$cacheKey]
    }
    
    try {
        if (-not (Get-Module -Name ActiveDirectory -ErrorAction SilentlyContinue)) {
            Import-Module ActiveDirectory -ErrorAction Stop
        }
        
        # Query parameters: if FQDNDomain is specified, use it as server
        $getADUserParams = @{
            Filter      = "SamAccountName -eq '$UserName'"
            Properties  = 'Enabled'
            ErrorAction = 'SilentlyContinue'
        }
        if ($ADConfig.FQDNDomain) {
            $getADUserParams.Server = $ADConfig.FQDNDomain
        }
        
        $adUser = Get-ADUser @getADUserParams
        
        if ($adUser) {
            $result = @{ Exists = $true; Enabled = $adUser.Enabled }
        }
        else {
            $result = @{ Exists = $false; Enabled = $false }
        }
    }
    catch {
        Write-Log "Warning: AD check error for '$UserName' in domain $($ADConfig.FQDNDomain): $_" -Level "WARN"
        $result = @{ Exists = $null; Enabled = $null }
    }
    
    $script:adCache[$cacheKey] = $result
    return $result
}

# ========== PROFILE DISCOVERY FUNCTION ==========
function Find-UserProfiles {
    param([string]$UserFolderPath)

    $userProfiles = [System.Collections.Generic.List[PSObject]]::new()

    function Get-ProfileType {
        param([string]$Path)
    
        $folderName = Split-Path $Path -Leaf
        Write-Log "Determining profile type for folder: $Path" -Level "DEBUG"
    
        # 1. Reset Citrix profile (folder name contains "upm_" but not "UPM_Profile")
        if ($folderName -match 'upm_' -and $folderName -notmatch 'UPM_Profile') {
            Write-Log "Detected reset Citrix profile by upm_ pattern in name" -Level "DEBUG"
            return [PSCustomObject]@{
                Path        = $Path
                ProfileType = 'ResetCitrixProfile'
                NtuserPath  = ''
                Details     = 'Reset Citrix profile (by upm_ pattern in folder name)'
            }
        }
    
        # Get child items once
        $children = Get-ChildItem -Path $Path -Force -ErrorAction SilentlyContinue
        if ($null -eq $children) {
            # Failed to get contents – maybe no permissions or folder unavailable
            Write-Log "Failed to get folder contents: $Path" -Level "DEBUG"
            return [PSCustomObject]@{
                Path        = $Path
                ProfileType = 'NotDefined'
                NtuserPath  = ''
                Details     = 'Could not read folder contents'
            }
        }
    
        # Create hashtables for quick lookup by name and type
        $dirs = @{}
        $files = @{}
        foreach ($child in $children) {
            if ($child.PSIsContainer) {
                $dirs[$child.Name] = $true
            }
            else {
                $files[$child.Name] = $true
            }
        }
    
        # 2. Citrix UPM indicators
        if ($dirs.ContainsKey('UPM_Profile') -or $dirs.ContainsKey('Pending') -or $files.ContainsKey('UPMSettings.ini')) {
            Write-Log "Detected UPM indicators" -Level "DEBUG"
            $ntuser = Join-Path $Path 'UPM_Profile\NTUSER.DAT'
            $details = "Found by UPM indicator"
            if (-not [System.IO.File]::Exists($ntuser)) {
                $details += " (NTUSER.DAT missing)"
            }
            return [PSCustomObject]@{
                Path        = $Path
                ProfileType = 'CitrixUPM'
                NtuserPath  = $ntuser
                Details     = $details
            }
        }
    
        # 3. Microsoft Roaming indicators
        if ($dirs.ContainsKey('AppData') -or $dirs.ContainsKey('Desktop') -or $files.ContainsKey('NTUSER.DAT') -or $files.ContainsKey('ntuser.ini')) {
            Write-Log "Detected Microsoft Roaming indicators" -Level "DEBUG"
            $ntuser = Join-Path $Path 'NTUSER.DAT'
            $details = "Found by MS indicator"
            if (-not [System.IO.File]::Exists($ntuser)) {
                $details += " (NTUSER.DAT missing)"
            }
            return [PSCustomObject]@{
                Path        = $Path
                ProfileType = 'MicrosoftRoaming'
                NtuserPath  = $ntuser
                Details     = $details
            }
        }
    
        # 4. Folder is empty
        if ($children.Count -eq 0) {
            Write-Log "Folder is empty" -Level "DEBUG"
            return [PSCustomObject]@{
                Path        = $Path
                ProfileType = 'NotDefined'
                NtuserPath  = 'empty'
                Details     = 'Empty profile'
            }
        }
    
        # 5. Unable to determine type
        Write-Log "Could not determine profile type by indicators" -Level "DEBUG"
        return [PSCustomObject]@{
            Path        = $Path
            ProfileType = 'NotDefined'
            NtuserPath  = 'empty'
            Details     = 'Missing NTUSER.DAT and UPM_Profile directory'
        }
    }

    # Analyze the root folder
    $rootResult = Get-ProfileType -Path $UserFolderPath

    if ($rootResult.ProfileType -ne 'NotDefined') {
        $userProfiles.Add($rootResult)
        return $userProfiles
    }

    $subFolders = Get-ChildItem -Path $UserFolderPath -Directory -ErrorAction SilentlyContinue

    if (-not $subFolders) {
        $userProfiles.Add($rootResult)
        return $userProfiles
    }

    foreach ($subFolder in $subFolders) {
        $result = Get-ProfileType -Path $subFolder.FullName
        $userProfiles.Add($result)
    }

    return $userProfiles
}

# ========== PROFILE HEALTH CHECK FUNCTION ==========
function Test-ProfileHealth {
    param(
        [string]$NtuserPath,
        [int]$MinNtuserSizeKB = 1
    )
    
    Write-Log "Checking health for path: $NtuserPath" -Level "DEBUG"
    
    if ([string]::IsNullOrEmpty($NtuserPath)) {
        Write-Log "NTUSER.DAT path is empty" -Level "DEBUG"
        return @{
            IsHealthy = $false
            Reason    = "NTUSER.DAT not specified"
            SizeKB    = 0
        }
    }
    
    # Check existence using .NET (works with long paths)
    if (-not [System.IO.File]::Exists($NtuserPath)) {
        Write-Log "NTUSER.DAT does not exist at path: $NtuserPath" -Level "DEBUG"
        return @{
            IsHealthy = $false
            Reason    = "NTUSER.DAT not found"
            SizeKB    = 0
        }
    }
    
    # Retry reading file properties via .NET FileInfo
    $maxRetries = 3
    $retryDelay = 200
    for ($i = 1; $i -le $maxRetries; $i++) {
        try {
            $fileInfo = [System.IO.FileInfo]::new($NtuserPath)
            $fileInfo.Refresh()
            if (-not $fileInfo.Exists) {
                throw "File does not exist (after Refresh)"
            }
            $sizeKB = [math]::Round($fileInfo.Length / 1KB, 2)
            
            if ($MinNtuserSizeKB -gt 0 -and $sizeKB -lt $MinNtuserSizeKB) {
                return @{
                    IsHealthy = $false
                    Reason    = "NTUSER.DAT too small ($sizeKB KB < $MinNtuserSizeKB KB)"
                    SizeKB    = $sizeKB
                }
            }
            
            # Also return last write time
            $lastWrite = $fileInfo.LastWriteTime
            return @{
                IsHealthy     = $true
                Reason        = "OK"
                SizeKB        = $sizeKB
                LastWriteTime = $lastWrite
            }
        }
        catch {
            Write-Log "Attempt $i of $maxRetries failed for path '$NtuserPath': $_" -Level "DEBUG"
            if ($i -eq $maxRetries) {
                Write-Log "Error accessing NTUSER.DAT after $maxRetries attempts: $_" -Level "ERROR"
                return @{
                    IsHealthy = $false
                    Reason    = "Error accessing NTUSER.DAT: $_"
                    SizeKB    = 0
                }
            }
            Start-Sleep -Milliseconds $retryDelay
        }
    }
}

# ========== PROFILE MOVE FUNCTION (with long path support) ==========
function Move-ProfileToQuarantine {
    param(
        [string]$SourcePath,
        [string]$DestinationPath, # full destination path (pre-built)
        [string]$userProfileType,
        [string]$Reason
    )

    try {
        if ($TestMode) {
            Write-Log "TEST: Profile [$userProfileType] would be moved: $SourcePath -> $DestinationPath (reason: $Reason)" -Level "INFO"
            return $true
        }

        Write-Log "Moving [$userProfileType]: $SourcePath -> $DestinationPath" -Level "INFO"

        # Create quarantine parent folder if it doesn't exist
        $quarantineRoot = Split-Path $DestinationPath -Parent
        if (-not (Test-Path $quarantineRoot)) {
            New-Item -Path $quarantineRoot -ItemType Directory -Force | Out-Null
            Write-Log "Created quarantine folder: $quarantineRoot" -Level "INFO"
        }

        # Build long path prefix for Move-Item
        $longSource = $SourcePath
        $longDest = $DestinationPath
        if ($SourcePath.StartsWith('\\')) {
            $longSource = '\\?\UNC\' + $SourcePath.Substring(2)
            $longDest = '\\?\UNC\' + $DestinationPath.Substring(2)
        }
        else {
            $longSource = '\\?\' + $SourcePath
            $longDest = '\\?\' + $DestinationPath
        }

        # Move the folder
        Move-Item -LiteralPath $longSource -Destination $longDest -Force -ErrorAction Stop

        # Verify that the source folder is gone and the destination exists
        if (Test-Path $DestinationPath) {
            # Update LastWriteTime to current time so quarantine period starts now
            try {
                (Get-Item -LiteralPath $DestinationPath).LastWriteTime = Get-Date
                Write-Log "Updated LastWriteTime of quarantine folder to current time" -Level "DEBUG"
            }
            catch {
                Write-Log "Warning: Failed to update LastWriteTime of $DestinationPath : $_" -Level "WARN"
            }

            Write-Log "Profile successfully moved" -Level "INFO"
            return $true
        }
        else {
            Write-Log "ERROR: After move, destination folder not found: $DestinationPath" -Level "ERROR"
            return $false
        }
    }
    catch {
        Write-Log "ERROR moving profile (exception): $_" -Level "ERROR"
        return $false
    }
}

# ========== PROFILE DELETE FUNCTION ==========
function Remove-Profile {
    param(
        [string]$UserProfilePath,
        [string]$UserProfileType,
        [string]$Reason
    )

    try {
        if ($TestMode) {
            Write-Log "TEST: Profile [$UserProfileType] would be deleted: $UserProfilePath (reason: $Reason)" -Level "INFO"
            return $true
        }

        Write-Log "Deleting [$UserProfileType]: $UserProfilePath (reason: $Reason)" -Level "INFO"

        # Check if the folder still exists
        if (-not (Test-Path -LiteralPath $UserProfilePath)) {
            Write-Log "Profile folder does not exist: $UserProfilePath" -Level "WARN"
            return $true
        }

        # Build long path prefix for Remove-Item to handle paths > 260 chars
        $longPath = $UserProfilePath
        if ($UserProfilePath.StartsWith('\\')) {
            # UNC path: \\server\share\...  ->  \\?\UNC\server\share\...
            $longPath = '\\?\UNC\' + $UserProfilePath.Substring(2)
        }
        else {
            # Local path: C:\Folder\...  ->  \\?\C:\Folder\...
            $longPath = '\\?\' + $UserProfilePath
        }

        # Use Remove-Item with -LiteralPath to avoid wildcard interpretation and support long paths
        Remove-Item -LiteralPath $longPath -Recurse -Force -ErrorAction Stop

        # Verify deletion
        if (-not (Test-Path -LiteralPath $UserProfilePath)) {
            Write-Log "Profile successfully deleted" -Level "INFO"
            return $true
        }
        else {
            Write-Log "ERROR: After deletion, folder still exists: $UserProfilePath" -Level "ERROR"
            return $false
        }
    }
    catch {
        Write-Log "ERROR deleting profile (exception): $_" -Level "ERROR"
        return $false
    }
}

# ========== FOLDER REDIRECTION MOVE FUNCTION (using robocopy for reliability) ==========
function Move-FolderRedirectionToQuarantine {
    param(
        [string]$SourcePath,
        [string]$DestinationPath,
        [string]$UserName,
        [string]$Reason
    )

    try {
        if ($TestMode) {
            Write-Log "TEST: Folder Redirection for [$UserName] would be moved: $SourcePath -> $DestinationPath (reason: $Reason)" -Level "INFO"
            return $true
        }

        Write-Log "Moving Folder Redirection [$UserName]: $SourcePath -> $DestinationPath" -Level "INFO"

        # Create quarantine parent folder if it doesn't exist
        $quarantineRoot = Split-Path $DestinationPath -Parent
        if (-not (Test-Path $quarantineRoot)) {
            New-Item -Path $quarantineRoot -ItemType Directory -Force | Out-Null
            Write-Log "Created quarantine folder: $quarantineRoot" -Level "INFO"
        }

        # Use robocopy to move the folder contents
        $robocopyArgs = @(
            "`"$SourcePath`"",
            "`"$DestinationPath`"",
            "/E",        # copy subdirectories, including empty ones
            "/MOV",      # move files (delete from source after copying)
            "/R:3",      # retry 3 times on failure
            "/W:5",      # wait 5 seconds between retries
            "/NP",       # no progress
            "/NJH",      # no job header
            "/NJS"       # no job summary
        )

        Write-Log "Executing: robocopy $($robocopyArgs -join ' ')" -Level "DEBUG"
        $result = & robocopy @robocopyArgs 2>&1
        $exitCode = $LASTEXITCODE

        # Robocopy exit codes 0-7 are considered success (some changes made)
        if ($exitCode -ge 8) {
            throw "Robocopy failed with exit code $exitCode. Output: $result"
        }

        # Clean up any remaining items in the source folder (hidden/system files)
        if (Test-Path $SourcePath) {
            try {
                Write-Log "Cleaning up remaining items in source folder: $SourcePath" -Level "DEBUG"
                # Build long path prefix for Remove-Item
                $longSourcePath = if ($SourcePath.StartsWith('\\')) {
                    '\\?\UNC\' + $SourcePath.Substring(2)
                }
                else {
                    '\\?\' + $SourcePath
                }
                Remove-Item -LiteralPath $longSourcePath -Recurse -Force -ErrorAction Stop
                Write-Log "Source folder successfully removed" -Level "DEBUG"
            }
            catch {
                Write-Log "Warning: Could not delete source folder '$SourcePath' after forced cleanup: $_" -Level "WARN"
                # Not a fatal error
            }
        }

        if (Test-Path $DestinationPath) {
            # Update LastWriteTime to current time for proper quarantine aging
            try {
                (Get-Item -LiteralPath $DestinationPath).LastWriteTime = Get-Date
                Write-Log "Updated LastWriteTime of quarantine folder" -Level "DEBUG"
            }
            catch {
                Write-Log "Warning: Failed to update LastWriteTime: $_" -Level "WARN"
            }

            Write-Log "Folder Redirection successfully moved" -Level "INFO"
            return $true
        }
        else {
            Write-Log "ERROR: After move, destination folder not found: $DestinationPath" -Level "ERROR"
            return $false
        }
    }
    catch {
        Write-Log "ERROR moving folder redirection (exception): $_" -Level "ERROR"
        return $false
    }
}

# ========== FOLDER REDIRECTION DELETE FUNCTION ==========
function Remove-FolderRedirection {
    param(
        [string]$UserFolderPath,
        [string]$UserName,
        [string]$Reason
    )

    try {
        if ($TestMode) {
            Write-Log "TEST: Folder Redirection for [$UserName] would be deleted: $UserFolderPath (reason: $Reason)" -Level "INFO"
            return $true
        }

        Write-Log "Deleting Folder Redirection [$UserName]: $UserFolderPath (reason: $Reason)" -Level "INFO"

        if (-not (Test-Path -LiteralPath $UserFolderPath)) {
            Write-Log "Folder does not exist: $UserFolderPath" -Level "WARN"
            return $true
        }

        # Build long path prefix
        $longPath = $UserFolderPath
        if ($UserFolderPath.StartsWith('\\')) {
            $longPath = '\\?\UNC\' + $UserFolderPath.Substring(2)
        }
        else {
            $longPath = '\\?\' + $UserFolderPath
        }

        Remove-Item -LiteralPath $longPath -Recurse -Force -ErrorAction Stop

        if (-not (Test-Path -LiteralPath $UserFolderPath)) {
            Write-Log "Folder Redirection successfully deleted" -Level "INFO"
            return $true
        }
        else {
            Write-Log "ERROR: After deletion, folder still exists: $UserFolderPath" -Level "ERROR"
            return $false
        }
    }
    catch {
        Write-Log "ERROR deleting folder redirection (exception): $_" -Level "ERROR"
        return $false
    }
}

# ========== QUARANTINE CLEANUP FUNCTION ==========
function Clear-OldQuarantine {
    param(
        [string]$QuarantinePath,
        [int]$QuarantineDays,
        [string]$SourceName   # Added to link action to source in report
    )
    
    if (-not (Test-Path $QuarantinePath)) {
        Write-Log "Quarantine folder does not exist: $QuarantinePath" -Level "DEBUG"
        return
    }
    
    $cutoffDate = (Get-Date).AddDays(-$QuarantineDays)
    $oldProfiles = Get-ChildItem -Path $QuarantinePath -Directory -ErrorAction SilentlyContinue |
    Where-Object { $_.LastWriteTime -lt $cutoffDate }
    
    if ($oldProfiles) {
        $count = @($oldProfiles).Count
        Write-Log "Found profiles in quarantine older than $QuarantineDays days: $count" -Level "INFO"
        
        foreach ($userProfile in $oldProfiles) {
            $profileName = $userProfile.Name
            $profileFullPath = $userProfile.FullName
            try {
                if ($TestMode) {
                    Write-Log "TEST: Profile would be deleted from quarantine: $profileName" -Level "INFO"
                    $status = "TEST"
                }
                else {
                    # Build long path prefix to support paths > 260 chars and special characters
                    $longPath = $profileFullPath
                    if ($profileFullPath.StartsWith('\\')) {
                        # UNC path: \\server\share\...  ->  \\?\UNC\server\share\...
                        $longPath = '\\?\UNC\' + $profileFullPath.Substring(2)
                    }
                    else {
                        # Local path: C:\Folder\...  ->  \\?\C:\Folder\...
                        $longPath = '\\?\' + $profileFullPath
                    }

                    # Use Remove-Item with -LiteralPath to avoid wildcard interpretation and support long paths
                    Remove-Item -LiteralPath $longPath -Recurse -Force -ErrorAction Stop

                    Write-Log "Deleted old profile from quarantine: $profileName" -Level "INFO"
                    $status = "SUCCESS"
                }
                
                # Add to execution details for HTML report
                $detail = [PSCustomObject]@{
                    Source       = $SourceName
                    UserName     = ""                       # Not applicable for quarantine cleanup
                    ProfileName  = $profileName
                    ProfileType  = "Quarantine"
                    Reason       = "OLD_QUARANTINE"
                    Details      = "Older than $QuarantineDays days"
                    Action       = "DELETE"
                    Destination  = "N/A"
                    Status       = $status
                    ErrorMessage = $null
                }
                $executionDetails.Add($detail)
                
                # Increment global stats counter
                $script:stats.OldQuarantineProfiles++
            }
            catch {
                Write-Log "Error deleting profile from quarantine: $profileName - $_" -Level "WARN"
                if (-not $TestMode) {
                    $detail = [PSCustomObject]@{
                        Source       = $SourceName
                        UserName     = ""
                        ProfileName  = $profileName
                        ProfileType  = "Quarantine"
                        Reason       = "OLD_QUARANTINE"
                        Details      = "Older than $QuarantineDays days"
                        Action       = "DELETE"
                        Destination  = "N/A"
                        Status       = "FAILED"
                        ErrorMessage = $_.Exception.Message
                    }
                    $executionDetails.Add($detail)
                    $script:stats.OldQuarantineProfiles++
                }
            }
        }
    }
    else {
        Write-Log "No old profiles found in quarantine" -Level "DEBUG"
    }
}

# ========== ORPHANED FOLDER REDIRECTION CLEANUP FUNCTION ==========
function Clear-OrphanedFolderRedirection {
    param(
        [string[]]$FolderRedirectionPaths,
        [string]$QuarantinePath,
        [bool]$EnableQuarantine,
        [string]$SourceName,
        [hashtable]$FoundUsers,      # usernames that have a corresponding profile folder
        [string]$ProcessOrphanedFR,  # "Disabled", "ReportOnly", or "Delete"
        [string[]]$ExcludePatterns   # patterns to exclude specific subfolders inside FR
    )

    if ($ProcessOrphanedFR -eq "Disabled") {
        Write-Log "Orphaned Folder Redirection processing is disabled for this source" -Level "INFO"
        return
    }

    Write-Log "Orphaned Folder Redirection mode: $ProcessOrphanedFR" -Level "INFO"

    foreach ($frPath in $FolderRedirectionPaths) {
        if (-not (Test-Path $frPath)) {
            Write-Log "Folder Redirection path not accessible: $frPath" -Level "WARN"
            continue
        }

        $frSubFolders = Get-ChildItem -Path $frPath -Directory -ErrorAction SilentlyContinue
        foreach ($frSubFolder in $frSubFolders) {
            $userName = $frSubFolder.Name

            # Skip if the subfolder name matches any FR exclude pattern
            $skip = $false
            if ($ExcludePatterns -and $ExcludePatterns.Count -gt 0) {
                foreach ($pattern in $ExcludePatterns) {
                    if ($userName -like $pattern) {
                        Write-Log "Skipping excluded orphaned FR folder: $($frSubFolder.FullName) (pattern: $pattern)" -Level "DEBUG"
                        $skip = $true
                        break
                    }
                }
            }
            if ($skip) { continue }

            if (-not $FoundUsers.ContainsKey($userName)) {
                Write-Log "Orphaned Folder Redirection found for user '$userName': $($frSubFolder.FullName)" -Level "INFO"

                if ($ProcessOrphanedFR -eq "ReportOnly") {
                    # Only report, no action
                    $frDetail = [PSCustomObject]@{
                        Source       = $SourceName
                        UserName     = $userName
                        ProfileName  = (Split-Path $frPath -Leaf) + "_" + $userName
                        ProfileType  = "FolderRedirection"
                        Reason       = "ORPHANED"
                        Details      = "No corresponding profile found (report only)"
                        Action       = "NONE"
                        Destination  = "N/A"
                        Status       = if ($TestMode) { "TEST" } else { "REPORTED" }
                        ErrorMessage = $null
                    }
                    $executionDetails.Add($frDetail)
                    $script:stats.OrphanedFRReportedCount++
                    continue
                }

                # Process (move or delete) orphaned FR
                if ($EnableQuarantine) {
                    $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
                    $safeUserName = $userName -replace '[\\/:*?"<>|]', '_'
                    $quarantineFolder = "FR_ORPHAN_${safeUserName}_${timestamp}"
                    $destinationPath = Join-Path $QuarantinePath $quarantineFolder

                    $frDetail = [PSCustomObject]@{
                        Source       = $SourceName
                        UserName     = $userName
                        ProfileName  = (Split-Path $frPath -Leaf) + "_" + $userName
                        ProfileType  = "FolderRedirection"
                        Reason       = "ORPHANED"
                        Details      = "No corresponding profile found"
                        Action       = "MOVE"
                        Destination  = $destinationPath
                        Status       = $null
                        ErrorMessage = $null
                    }

                    $script:stats.FolderRedirectionToMoveCount++
                    $script:stats.OrphanedFRProcessedCount++
                    if ($TestMode) {
                        $frDetail.Status = "TEST"
                        Write-Log "TEST: Orphaned FR would be moved: $($frSubFolder.FullName) -> $destinationPath" -Level "INFO"
                    }
                    else {
                        $moved = Move-FolderRedirectionToQuarantine -SourcePath $frSubFolder.FullName -DestinationPath $destinationPath -UserName $userName -Reason "ORPHANED"
                        if ($moved) {
                            $frDetail.Status = "SUCCESS"
                            $script:stats.FolderRedirectionMovedSuccess++
                        }
                        else {
                            $frDetail.Status = "FAILED"
                            $frDetail.ErrorMessage = "Move error"
                            $script:stats.FolderRedirectionMoveFailed++
                        }
                    }
                    $executionDetails.Add($frDetail)
                }
                else {
                    $frDetail = [PSCustomObject]@{
                        Source       = $SourceName
                        UserName     = $userName
                        ProfileName  = (Split-Path $frPath -Leaf) + "_" + $userName
                        ProfileType  = "FolderRedirection"
                        Reason       = "ORPHANED"
                        Details      = "No corresponding profile found"
                        Action       = "DELETE"
                        Destination  = "N/A"
                        Status       = $null
                        ErrorMessage = $null
                    }

                    $script:stats.FolderRedirectionToDeleteCount++
                    $script:stats.OrphanedFRProcessedCount++
                    if ($TestMode) {
                        $frDetail.Status = "TEST"
                        Write-Log "TEST: Orphaned FR would be deleted: $($frSubFolder.FullName)" -Level "INFO"
                    }
                    else {
                        $removed = Remove-FolderRedirection -UserFolderPath $frSubFolder.FullName -UserName $userName -Reason "ORPHANED"
                        if ($removed) {
                            $frDetail.Status = "SUCCESS"
                            $script:stats.FolderRedirectionDeletedSuccess++
                        }
                        else {
                            $frDetail.Status = "FAILED"
                            $frDetail.ErrorMessage = "Delete error"
                            $script:stats.FolderRedirectionDeleteFailed++
                        }
                    }
                    $executionDetails.Add($frDetail)
                }
            }
        }
    }
}

# ========== EMPTY FOLDER CLEANUP FUNCTION (WITH EXCLUSIONS) ==========
function Remove-EmptyFolders {
    param(
        [string]$FolderPath, # root folder (ProfileRoot)
        [string[]]$ExcludePatterns    # patterns for excluded folder names
    )
    
    # Get all folders directly under the root
    $subFolders = Get-ChildItem -Path $FolderPath -Directory -Force -ErrorAction SilentlyContinue
    
    # Apply exclusion patterns
    if ($ExcludePatterns -and $ExcludePatterns.Count -gt 0) {
        $subFolders = $subFolders | Where-Object {
            $name = $_.Name
            $excluded = $false
            foreach ($pattern in $ExcludePatterns) {
                if ($name -like $pattern) {
                    $excluded = $true
                    break
                }
            }
            -not $excluded
        }
    }
    
    foreach ($subFolder in $subFolders) {
        # Check if the folder is empty (including hidden items)
        $items = Get-ChildItem -Path $subFolder.FullName -Force -ErrorAction SilentlyContinue
        if (-not $items) {
            if ($TestMode) {
                Write-Log "TEST: Empty folder would be deleted: $($subFolder.FullName)" -Level "INFO"
                $script:stats.EmptyFolders++
            }
            else {
                try {
                    Remove-Item -Path $subFolder.FullName -Force -ErrorAction SilentlyContinue
                    Write-Log "Deleted empty folder: $($subFolder.FullName)" -Level "INFO"
                    $script:stats.EmptyFolders++
                }
                catch {
                    Write-Log "Could not delete empty folder $($subFolder.FullName): $_" -Level "WARN"
                }
            }
        }
    }
}

# ========== HELPER FUNCTIONS ==========
function Format-Count {
    param([int]$Success, [int]$Total)
    if ($Total -eq 0) { return "0" }
    return "$Success / $Total"
}

# ========== MAIN SCRIPT ==========
Write-Log "=" * 70 -Level "INFO"
if ($TestMode) {
    Write-Log "RUNNING SCRIPT IN TEST MODE" -Level "INFO"
}
else {
    Write-Log "RUNNING SCRIPT IN REAL MODE" -Level "INFO"
}
Write-Log "Date: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')" -Level "INFO"
Write-Log "=" * 70 -Level "INFO"

# Statistics (extended)
$stats = @{
    TotalSources                    = @($userProfileSources).Count
    SourcesProcessed                = 0
    TotalProfilesFound              = 0
    CitrixProfilesFound             = 0
    RoamingProfilesFound            = 0
    ResetProfilesFound              = 0
    NotDefinedProfilesFound         = 0
    ProfilesToMoveCount             = 0
    ProfilesToDeleteCount           = 0
    ProfilesToMoveByAge             = 0
    ProfilesToMoveByAD              = 0
    ProfilesToMoveByCorruption      = 0
    ResetProfilesMoved              = 0
    ProfilesMovedSuccess            = 0
    ProfilesMoveFailed              = 0
    ProfilesDeletedSuccess          = 0
    ProfilesDeleteFailed            = 0
    FolderRedirectionMovedSuccess   = 0
    FolderRedirectionMoveFailed     = 0
    FolderRedirectionDeletedSuccess = 0
    FolderRedirectionDeleteFailed   = 0
    FolderRedirectionToMoveCount    = 0
    FolderRedirectionToDeleteCount  = 0
    OrphanedFRProcessedCount        = 0
    OrphanedFRReportedCount         = 0
    OldQuarantineProfiles           = 0
    EmptyFolders                    = 0
}

# Execution details for HTML report
$executionDetails = [System.Collections.Generic.List[PSObject]]::new()
# Detailed statistics per source for final report
$sourceReports = [System.Collections.Generic.List[PSObject]]::new()

foreach ($source in $userProfileSources) {
    Write-Log "" -Level "INFO"
    Write-Log "Source: $($source.Name)" -Level "INFO"
    Write-Log "Path: $($source.ProfileRoot)" -Level "INFO"
    Write-Log "Pattern: $($source.PatternProfile)" -Level "INFO"
    Write-Log "DaysToDelete: $($source.DaysToDelete)" -Level "INFO"
    Write-Log "Quarantine: $($source.QuarantinePath)" -Level "INFO"
    Write-Log "QuarantineDays: $($source.QuarantineDays)" -Level "INFO"
    # Option to enable quarantine (default true)
    $enableQuarantine = if ($source.PSObject.Properties.Name -contains 'EnableQuarantine') { [bool]$source.EnableQuarantine } else { $true }
    Write-Log "Quarantine enabled: $enableQuarantine" -Level "INFO"
    if ($source.ExcludeFolders) {
        Write-Log "Excluded folders: $($source.ExcludeFolders -join ', ')" -Level "INFO"
    }

    # --- Determine AD settings for this source ---
    if ($source.PSObject.Properties.Name -contains 'ActiveDirectory' -and $source.ActiveDirectory) {
        $sourceAD = $source.ActiveDirectory
        if (-not ($sourceAD.PSObject.Properties.Name -contains 'Enabled')) {
            $sourceAD.Enabled = $GlobalADSettings.Enabled
        }
        if (-not $sourceAD.FQDNDomain) { $sourceAD.FQDNDomain = $GlobalADSettings.FQDNDomain }
        if (-not $sourceAD.NetBIOSDomain) { $sourceAD.NetBIOSDomain = $GlobalADSettings.NetBIOSDomain }
        if (-not $sourceAD.ExcludeUsers) { $sourceAD.ExcludeUsers = $GlobalADSettings.ExcludeUsers }
    }
    else {
        $sourceAD = $GlobalADSettings
    }

    # Escape domains for regular expressions (username cleanup)
    if ($sourceAD.Enabled -and $sourceAD.FQDNDomain -and $sourceAD.NetBIOSDomain) {
        $sourceNetBios = [regex]::Escape($sourceAD.NetBIOSDomain)
        $sourceFQDN = [regex]::Escape($sourceAD.FQDNDomain)
    }
    else {
        $sourceNetBios = [regex]::Escape($GlobalADSettings.NetBIOSDomain)
        $sourceFQDN = [regex]::Escape($GlobalADSettings.FQDNDomain)
    }

    # Determine orphaned FR processing mode
    $orphanFRMode = "Disabled"  # default
    if ($source.PSObject.Properties.Name -contains 'ProcessOrphanedFR') {
        $value = $source.ProcessOrphanedFR
        if ($value -is [bool]) {
            $orphanFRMode = if ($value) { "Delete" } else { "Disabled" }
        }
        elseif ($value -is [string]) {
            $orphanFRMode = $value
        }
    }

    # Initialize source-level stats
    $sourceStats = @{
        UsersProcessed         = 0
        ProfilesFound          = 0
        CitrixProfiles         = 0
        RoamingProfiles        = 0
        ResetProfiles          = 0
        NotDefinedProfiles     = 0
        ProfilesToMove         = 0
        ProfilesToDelete       = 0
        ProfilesMovedSuccess   = 0
        ProfilesMoveFailed     = 0
        ProfilesDeletedSuccess = 0
        ProfilesDeleteFailed   = 0
        FRToMove               = 0
        FRToDelete             = 0
        FRMovedSuccess         = 0
        FRMoveFailed           = 0
        FRDeletedSuccess       = 0
        FRDeleteFailed         = 0
        OrphanedFRProcessed    = 0
        OrphanedFRReported     = 0
        QuarantineCleaned      = 0
    }

    try {
        if (-not (Test-Path $source.ProfileRoot)) {
            Write-Log "ERROR: Network path is not accessible" -Level "ERROR"
            continue
        }
        
        $userFolders = Get-ChildItem -Path $source.ProfileRoot -Directory -ErrorAction SilentlyContinue |
        Where-Object { ($_.Name -like $source.PatternProfile) -and ($_.FullName -notlike "$($source.QuarantinePath)*") }
        
        # Build list of full paths to exclude from profile scanning (folder redirection roots)
        $excludeFullPaths = @()
        if ($source.PSObject.Properties.Name -contains 'FolderRedirectionPaths' -and $source.FolderRedirectionPaths) {
            $excludeFullPaths += $source.FolderRedirectionPaths
        }

        # Get FR exclude patterns
        $frExcludePatterns = @()
        if ($source.PSObject.Properties.Name -contains 'FolderRedirectionExcludePatterns' -and $source.FolderRedirectionExcludePatterns) {
            $frExcludePatterns = $source.FolderRedirectionExcludePatterns
        }

        # Filter out user folders that are located under any folder redirection root,
        # and also exclude specific subfolders inside FR paths based on patterns
        if ($excludeFullPaths.Count -gt 0) {
            $userFolders = $userFolders | Where-Object {
                $folderPath = $_.FullName
                $excluded = $false
                foreach ($exPath in $excludeFullPaths) {
                    # Check if folder is exactly the FR root or a subfolder under it
                    if ($folderPath -eq $exPath -or $folderPath.StartsWith("$exPath\", 'OrdinalIgnoreCase')) {
                        # If it's inside an FR path, check against FR exclude patterns
                        if ($frExcludePatterns.Count -gt 0) {
                            $relativePath = $folderPath.Substring($exPath.Length).TrimStart('\')
                            if ($relativePath -eq '') {
                                # This is the FR root itself – always exclude it from profile scanning
                                $excluded = $true
                            }
                            else {
                                # Check if the relative path (or any part) matches any exclusion pattern
                                foreach ($pattern in $frExcludePatterns) {
                                    if ($relativePath -like $pattern) {
                                        $excluded = $true
                                        Write-Log "Excluded folder (FR exclude pattern '$pattern'): $folderPath" -Level "DEBUG"
                                        break
                                    }
                                }
                            }
                        }
                        else {
                            # No FR exclude patterns, exclude the entire FR root and all subfolders
                            $excluded = $true
                        }
                        if (-not $excluded) {
                            Write-Log "Not excluded (inside FR but not matched by patterns): $folderPath" -Level "DEBUG"
                        }
                        break
                    }
                }
                -not $excluded
            }
        }

        # Apply ExcludeFolders exclusions
        if ($source.ExcludeFolders -and $source.ExcludeFolders.Count -gt 0) {
            $userFolders = $userFolders | Where-Object {
                $name = $_.Name
                $excluded = $false
                foreach ($pattern in $source.ExcludeFolders) {
                    if ($name -like $pattern) {
                        $excluded = $true
                        break
                    }
                }
                -not $excluded
            }
            Write-Log "After exclusions, remaining folders: $($userFolders.Count)" -Level "DEBUG"
        }
        
        if (-not $userFolders) {
            Write-Log "No user folders found (or all excluded)" -Level "INFO"
            # Even if no user folders, we can still check for orphaned FR
            if ($orphanFRMode -ne "Disabled" -and $source.PSObject.Properties.Name -contains 'FolderRedirectionPaths') {
                $foundUsers = @{}  # empty hashtable
                Clear-OrphanedFolderRedirection -FolderRedirectionPaths $source.FolderRedirectionPaths `
                    -QuarantinePath $source.QuarantinePath `
                    -EnableQuarantine $enableQuarantine `
                    -SourceName $source.Name `
                    -FoundUsers $foundUsers `
                    -ProcessOrphanedFR $orphanFRMode `
                    -ExcludePatterns $frExcludePatterns
            }
            continue
        }
        
        Write-Log "Found user folders: $($userFolders.Count)" -Level "INFO"
        
        # Collect usernames for orphaned FR detection (normalized)
        $foundUsers = @{}
        
        foreach ($userFolder in $userFolders) {
            # Extract username
            $rawName = $userFolder.Name
            $userName = $rawName `
                -replace '\.v\d+$', '' `
                -replace '^upm_', '' `
                -replace "^$sourceNetBios\.", '' `
                -replace "^$sourceFQDN\.", '' `
                -replace "\.$sourceNetBios$", '' `
                -replace "\.$sourceFQDN$", ''

            $foundUsers[$userName] = $true

            $sourceStats.UsersProcessed++
            
            if ($GeneralSettings.DetailedLogging -or $TestMode) {
                Write-Log "  User: $userName (original folder: $rawName)" -Level "DEBUG"
            }
            
            # Find profiles inside the user folder
            $userProfiles = Find-UserProfiles -UserFolderPath $userFolder.FullName
            
            if (-not $userProfiles) {
                if ($GeneralSettings.DetailedLogging -or $TestMode) {
                    Write-Log "    No profiles found" -Level "DEBUG"
                }
                continue
            }
            
            Write-Log "    Profiles found for user: $(@($userProfiles).Count)" -Level "DEBUG"
            $sourceStats.ProfilesFound += @($userProfiles).Count

            # Flag to indicate if any action was taken for this user's profiles
            $userHadAction = $false
            
            foreach ($userProfile in $userProfiles) {
                $userProfileName = Split-Path $userProfile.Path -Leaf
                $userProfileType = $userProfile.ProfileType
                
                switch ($userProfileType) {
                    "CitrixUPM" {
                        $sourceStats.CitrixProfiles++
                        $stats.CitrixProfilesFound++
                    }
                    "MicrosoftRoaming" {
                        $sourceStats.RoamingProfiles++
                        $stats.RoamingProfilesFound++
                    }
                    "ResetCitrixProfile" {
                        $sourceStats.ResetProfiles++
                        $stats.ResetProfilesFound++
                    }
                    "NotDefined" {
                        $sourceStats.NotDefinedProfiles++
                        $stats.NotDefinedProfilesFound++
                    }
                }
                
                $stats.TotalProfilesFound++
                
                # --- Determine reason for move/delete ---
                $moveReason = $null
                $details = ""

                # Handle reset profiles (no AD check needed)
                if ($userProfileType -eq "ResetCitrixProfile") {
                    $moveReason = "RESET_PROFILE"
                    $details = $userProfile.Details
                }
                else {
                    # For normal profiles, first check health
                    $health = Test-ProfileHealth -NtuserPath $userProfile.NtuserPath -MinNtuserSizeKB 1

                    if (-not $health.IsHealthy) {
                        $moveReason = "PROFILE_CORRUPTED"
                        $details = $health.Reason
                    }
                    else {
                        # Check profile age (last write of NTUSER.DAT)
                        $lastWrite = $health.LastWriteTime
                        $cutoffDate = (Get-Date).AddDays(-$source.DaysToDelete)
                        $daysOld = [math]::Round(((Get-Date) - $lastWrite).TotalDays, 1)

                        if ($lastWrite -lt $cutoffDate) {
                            $moveReason = "PROFILE_OLD"
                            $details = "Inactive $daysOld days (threshold: $($source.DaysToDelete) days)"
                        }
                        else {
                            # Profile is not old – check AD
                            if ($sourceAD.Enabled) {
                                $userStatus = Get-UserADStatus -UserName $userName -ADConfig $sourceAD
                            }
                            else {
                                $userStatus = @{ Exists = $true; Enabled = $true }
                            }

                            if ($userStatus.Exists -eq $false) {
                                $moveReason = "USER_NOT_FOUND"
                                $details = ""
                            }
                            elseif ($userStatus.Enabled -eq $false) {
                                $moveReason = "USER_DISABLED"
                                $details = ""
                            }
                            # If user is active – no reason, profile remains
                        }
                    }
                }
                
                # If there is a reason – profile is subject to processing
                if ($moveReason) {
                    $userHadAction = $true
                    
                    # Determine action: MOVE (quarantine) or DELETE
                    $action = if ($enableQuarantine) { "MOVE" } else { "DELETE" }
                    
                    # For MOVE, build destination path
                    $destinationPath = $null
                    if ($action -eq "MOVE") {
                        $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
                        $safeProfileName = $userProfileName -replace '[\\/:*?"<>|]', '_'
                        
                        if ($userProfileType -eq "ResetCitrixProfile") {
                            $quarantineFolder = "RESET_${safeProfileName}_${timestamp}"
                        }
                        else {
                            $quarantineFolder = "${safeProfileName}_${userProfileType}_${timestamp}_${moveReason}"
                        }
                        
                        $destinationPath = Join-Path $source.QuarantinePath $quarantineFolder
                    }
                    
                    # Create detail object
                    $detail = [PSCustomObject]@{
                        Source       = $source.Name
                        UserName     = $userName
                        ProfileName  = $userProfileName
                        ProfileType  = $userProfileType
                        Reason       = $moveReason
                        Details      = $details
                        Action       = $action
                        Destination  = if ($destinationPath) { $destinationPath } else { "N/A" }
                        Status       = $null
                        ErrorMessage = $null
                    }
                    
                    # Update statistics based on action
                    if ($action -eq "MOVE") {
                        $stats.ProfilesToMoveCount++
                        $sourceStats.ProfilesToMove++
                    }
                    else {
                        $stats.ProfilesToDeleteCount++
                        $sourceStats.ProfilesToDelete++
                    }
                    
                    # Additional counters by reason
                    switch ($moveReason) {
                        "PROFILE_OLD" { $stats.ProfilesToMoveByAge++ }
                        "PROFILE_CORRUPTED" { $stats.ProfilesToMoveByCorruption++ }
                        "RESET_PROFILE" { $stats.ResetProfilesMoved++ }
                        default { $stats.ProfilesToMoveByAD++ }
                    }
                    
                    if ($TestMode) {
                        # Test mode: only log, do not execute action
                        Write-Log "TEST: Profile [$userProfileType] would be $(if($action -eq 'MOVE'){'moved'}else{'deleted'}): $($userProfile.Path) $(if($action -eq 'MOVE'){'-> '+$destinationPath}) (reason: $moveReason)" -Level "INFO"
                        $detail.Status = "TEST"
                        # For test mode, call the function that logs but does nothing (Move-ProfileToQuarantine in test mode returns true)
                        if ($action -eq "MOVE") {
                            $null = Move-ProfileToQuarantine -SourcePath $userProfile.Path -DestinationPath $destinationPath -ProfileType $userProfileType -Reason $moveReason
                        }
                        else {
                            $null = Remove-Profile -UserProfilePath $userProfile.Path -UserProfileType $userProfileType -Reason $moveReason
                        }
                    }
                    else {
                        # Real mode: execute action
                        if ($action -eq "MOVE") {
                            $moved = Move-ProfileToQuarantine -SourcePath $userProfile.Path -DestinationPath $destinationPath -ProfileType $userProfileType -Reason $moveReason
                            if ($moved) {
                                $detail.Status = "SUCCESS"
                                $stats.ProfilesMovedSuccess++
                            }
                            else {
                                $detail.Status = "FAILED"
                                $detail.ErrorMessage = "Move error (see log for details)"
                                $stats.ProfilesMoveFailed++
                            }
                        }
                        else {
                            $removed = Remove-Profile -UserProfilePath $userProfile.Path -UserProfileType $userProfileType -Reason $moveReason
                            if ($removed) {
                                $detail.Status = "SUCCESS"
                                $stats.ProfilesDeletedSuccess++
                            }
                            else {
                                $detail.Status = "FAILED"
                                $detail.ErrorMessage = "Delete error (see log for details)"
                                $stats.ProfilesDeleteFailed++
                            }
                        }
                    }
                    
                    # Add detail to the global list
                    $executionDetails.Add($detail)
                }
            } # foreach profile

            # ========== FOLDER REDIRECTION HANDLING (once per user, after processing all profiles) ==========
            if ($userHadAction -and $source.PSObject.Properties.Name -contains 'FolderRedirectionPaths' -and $source.FolderRedirectionPaths) {
                foreach ($frPath in $source.FolderRedirectionPaths) {
                    $frUserFolder = Join-Path $frPath $userName
                    if (Test-Path -LiteralPath $frUserFolder) {
                        if ($enableQuarantine) {
                            $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
                            $safeUserName = $userName -replace '[\\/:*?"<>|]', '_'
                            $quarantineFolder = "FR_${safeUserName}_${timestamp}"
                            $destinationPath = Join-Path $source.QuarantinePath $quarantineFolder

                            $frDetail = [PSCustomObject]@{
                                Source       = $source.Name
                                UserName     = $userName
                                ProfileName  = (Split-Path $frPath -Leaf) + "_" + $userName
                                ProfileType  = "FolderRedirection"
                                Reason       = "USER_PROFILE_REMOVED"
                                Details      = "Associated profile was removed"
                                Action       = "MOVE"
                                Destination  = $destinationPath
                                Status       = $null
                                ErrorMessage = $null
                            }

                            $stats.FolderRedirectionToMoveCount++
                            $sourceStats.FRToMove++
                            if ($TestMode) {
                                $frDetail.Status = "TEST"
                                $null = Move-FolderRedirectionToQuarantine -SourcePath $frUserFolder -DestinationPath $destinationPath -UserName $userName -Reason "USER_PROFILE_REMOVED"
                            }
                            else {
                                $moved = Move-FolderRedirectionToQuarantine -SourcePath $frUserFolder -DestinationPath $destinationPath -UserName $userName -Reason "USER_PROFILE_REMOVED"
                                if ($moved) {
                                    $frDetail.Status = "SUCCESS"
                                    $stats.FolderRedirectionMovedSuccess++
                                }
                                else {
                                    $frDetail.Status = "FAILED"
                                    $frDetail.ErrorMessage = "Move error"
                                    $stats.FolderRedirectionMoveFailed++
                                }
                            }
                            $executionDetails.Add($frDetail)
                        }
                        else {
                            $frDetail = [PSCustomObject]@{
                                Source       = $source.Name
                                UserName     = $userName
                                ProfileName  = (Split-Path $frPath -Leaf) + "_" + $userName
                                ProfileType  = "FolderRedirection"
                                Reason       = "USER_PROFILE_REMOVED"
                                Details      = "Associated profile was removed"
                                Action       = "DELETE"
                                Destination  = "N/A"
                                Status       = $null
                                ErrorMessage = $null
                            }

                            $stats.FolderRedirectionToDeleteCount++
                            $sourceStats.FRToDelete++
                            if ($TestMode) {
                                $frDetail.Status = "TEST"
                                $null = Remove-FolderRedirection -UserFolderPath $frUserFolder -UserName $userName -Reason "USER_PROFILE_REMOVED"
                            }
                            else {
                                $removed = Remove-FolderRedirection -UserFolderPath $frUserFolder -UserName $userName -Reason "USER_PROFILE_REMOVED"
                                if ($removed) {
                                    $frDetail.Status = "SUCCESS"
                                    $stats.FolderRedirectionDeletedSuccess++
                                }
                                else {
                                    $frDetail.Status = "FAILED"
                                    $frDetail.ErrorMessage = "Delete error"
                                    $stats.FolderRedirectionDeleteFailed++
                                }
                            }
                            $executionDetails.Add($frDetail)
                        }
                    }
                    else {
                        Write-Log "Folder Redirection path not found for user $userName at $frUserFolder" -Level "DEBUG"
                    }
                }
            }
        } # foreach userFolder
        
        # Process orphaned folder redirection (users with FR folders but no profile)
        if ($orphanFRMode -ne "Disabled" -and $source.PSObject.Properties.Name -contains 'FolderRedirectionPaths' -and $source.FolderRedirectionPaths) {
            Clear-OrphanedFolderRedirection -FolderRedirectionPaths $source.FolderRedirectionPaths `
                -QuarantinePath $source.QuarantinePath `
                -EnableQuarantine $enableQuarantine `
                -SourceName $source.Name `
                -FoundUsers $foundUsers `
                -ProcessOrphanedFR $orphanFRMode `
                -ExcludePatterns $frExcludePatterns
        }

        # Clean up old quarantine (only if quarantine is enabled and path exists)
        if ($enableQuarantine) {
            Clear-OldQuarantine -QuarantinePath $source.QuarantinePath -QuarantineDays $source.QuarantineDays -SourceName $source.Name
        }
        
        # Delete empty folders if enabled
        if ($source.PSObject.Properties.Name -contains 'EnableEmptyFolderCleanup') {
            $enableCleanup = [bool]$source.EnableEmptyFolderCleanup
        }
        else {
            $enableCleanup = $true   # enabled by default
        }

        if ($enableCleanup) {
            $excludePatterns = @()
            if ($source.PSObject.Properties.Name -contains 'EmptyFolderExcludePatterns' -and $source.EmptyFolderExcludePatterns) {
                $excludePatterns = $source.EmptyFolderExcludePatterns
            }
            elseif ($source.ExcludeFolders) {
                $excludePatterns = $source.ExcludeFolders
            }
    
            Remove-EmptyFolders -FolderPath $source.ProfileRoot -ExcludePatterns $excludePatterns
        }
        
        # Collect detailed statistics from executionDetails for this source
        $sourceName = $source.Name
        $sourceStats.ProfilesMovedSuccess = ($executionDetails | Where-Object { $_.Source -eq $sourceName -and $_.Action -eq "MOVE" -and $_.ProfileType -ne "FolderRedirection" -and $_.Status -eq "SUCCESS" }).Count
        $sourceStats.ProfilesMoveFailed = ($executionDetails | Where-Object { $_.Source -eq $sourceName -and $_.Action -eq "MOVE" -and $_.ProfileType -ne "FolderRedirection" -and $_.Status -eq "FAILED" }).Count
        $sourceStats.ProfilesDeletedSuccess = ($executionDetails | Where-Object { $_.Source -eq $sourceName -and $_.Action -eq "DELETE" -and $_.ProfileType -ne "FolderRedirection" -and $_.Status -eq "SUCCESS" }).Count
        $sourceStats.ProfilesDeleteFailed = ($executionDetails | Where-Object { $_.Source -eq $sourceName -and $_.Action -eq "DELETE" -and $_.ProfileType -ne "FolderRedirection" -and $_.Status -eq "FAILED" }).Count
        $sourceStats.FRMovedSuccess = ($executionDetails | Where-Object { $_.Source -eq $sourceName -and $_.Action -eq "MOVE" -and $_.ProfileType -eq "FolderRedirection" -and $_.Reason -ne "ORPHANED" -and $_.Status -eq "SUCCESS" }).Count
        $sourceStats.FRMoveFailed = ($executionDetails | Where-Object { $_.Source -eq $sourceName -and $_.Action -eq "MOVE" -and $_.ProfileType -eq "FolderRedirection" -and $_.Reason -ne "ORPHANED" -and $_.Status -eq "FAILED" }).Count
        $sourceStats.FRDeletedSuccess = ($executionDetails | Where-Object { $_.Source -eq $sourceName -and $_.Action -eq "DELETE" -and $_.ProfileType -eq "FolderRedirection" -and $_.Reason -ne "ORPHANED" -and $_.Status -eq "SUCCESS" }).Count
        $sourceStats.FRDeleteFailed = ($executionDetails | Where-Object { $_.Source -eq $sourceName -and $_.Action -eq "DELETE" -and $_.ProfileType -eq "FolderRedirection" -and $_.Reason -ne "ORPHANED" -and $_.Status -eq "FAILED" }).Count
        $sourceStats.OrphanedFRProcessed = ($executionDetails | Where-Object { $_.Source -eq $sourceName -and $_.ProfileType -eq "FolderRedirection" -and $_.Reason -eq "ORPHANED" -and $_.Status -eq "SUCCESS" }).Count
        $sourceStats.OrphanedFRReported = ($executionDetails | Where-Object { $_.Source -eq $sourceName -and $_.ProfileType -eq "FolderRedirection" -and $_.Reason -eq "ORPHANED" -and $_.Status -in @("REPORTED", "TEST") }).Count
        $sourceStats.QuarantineCleaned = ($executionDetails | Where-Object { $_.Source -eq $sourceName -and $_.ProfileType -eq "Quarantine" -and $_.Status -in @("SUCCESS", "TEST") }).Count

        # Add to source reports collection
        $sourceReport = [PSCustomObject]@{
            SourceName             = $sourceName
            UsersProcessed         = $sourceStats.UsersProcessed
            ProfilesFound          = $sourceStats.ProfilesFound
            CitrixProfiles         = $sourceStats.CitrixProfiles
            RoamingProfiles        = $sourceStats.RoamingProfiles
            ResetProfiles          = $sourceStats.ResetProfiles
            NotDefinedProfiles     = $sourceStats.NotDefinedProfiles
            ProfilesToMove         = $sourceStats.ProfilesToMove
            ProfilesToDelete       = $sourceStats.ProfilesToDelete
            ProfilesMovedSuccess   = $sourceStats.ProfilesMovedSuccess
            ProfilesMoveFailed     = $sourceStats.ProfilesMoveFailed
            ProfilesDeletedSuccess = $sourceStats.ProfilesDeletedSuccess
            ProfilesDeleteFailed   = $sourceStats.ProfilesDeleteFailed
            FRToMove               = $sourceStats.FRToMove
            FRToDelete             = $sourceStats.FRToDelete
            FRMovedSuccess         = $sourceStats.FRMovedSuccess
            FRMoveFailed           = $sourceStats.FRMoveFailed
            FRDeletedSuccess       = $sourceStats.FRDeletedSuccess
            FRDeleteFailed         = $sourceStats.FRDeleteFailed
            OrphanedFRProcessed    = $sourceStats.OrphanedFRProcessed
            OrphanedFRReported     = $sourceStats.OrphanedFRReported
            QuarantineCleaned      = $sourceStats.QuarantineCleaned
        }
        $sourceReports.Add($sourceReport)

        Write-Log "" -Level "INFO"
        Write-Log "Statistics for source '$sourceName':" -Level "INFO"
        Write-Log "  Users: $($sourceStats.UsersProcessed)" -Level "INFO"
        Write-Log "  Profiles found: $($sourceStats.ProfilesFound)" -Level "INFO"
        Write-Log "    - Citrix UPM: $($sourceStats.CitrixProfiles)" -Level "INFO"
        Write-Log "    - Microsoft Roaming: $($sourceStats.RoamingProfiles)" -Level "INFO"
        Write-Log "    - Reset Citrix: $($sourceStats.ResetProfiles)" -Level "INFO"
        Write-Log "    - Undefined/empty profiles: $($sourceStats.NotDefinedProfiles)" -Level "INFO"
        
        if ($TestMode) {
            Write-Log "  Profiles to process (TEST): move $($sourceStats.ProfilesToMove), delete $($sourceStats.ProfilesToDelete)" -Level "INFO"
        }
        else {
            Write-Log "  Profiles processed: moved $(Format-Count -Success $sourceStats.ProfilesMovedSuccess -Total $sourceStats.ProfilesToMove) (err $($sourceStats.ProfilesMoveFailed)), deleted $(Format-Count -Success $sourceStats.ProfilesDeletedSuccess -Total $sourceStats.ProfilesToDelete) (err $($sourceStats.ProfilesDeleteFailed))" -Level "INFO"
        }
        
        $stats.SourcesProcessed++
    }
    catch {
        Write-Log "ERROR processing source '$($source.Name)': $_" -Level "ERROR"
    }
}

# Delete old logs
try {
    $logCutoffDate = (Get-Date).AddDays(-$GeneralSettings.LogRetentionDays)
    $oldLogs = Get-ChildItem -Path $LogDir -File -Filter "ProfileCleanup_*.log" -ErrorAction SilentlyContinue |
    Where-Object { $_.CreationTime -lt $logCutoffDate }
    
    if ($oldLogs) {
        foreach ($log in $oldLogs) {
            if (-not $TestMode) {
                Remove-Item -Path $log.FullName -Force -ErrorAction SilentlyContinue
            }
        }
        Write-Log "Old logs found for deletion: $($oldLogs.Count)" -Level "INFO"
    }
}
catch { }

# ========== FINAL STATISTICS ==========
Write-Log "" -Level "INFO"
Write-Log "=" * 70 -Level "INFO"
if ($TestMode) {
    Write-Log "TEST MODE - CHECK RESULTS" -Level "INFO"
}
else {
    Write-Log "REAL MODE - EXECUTION RESULTS" -Level "INFO"
}
Write-Log "=" * 70 -Level "INFO"

# Overall summary
Write-Log "TOTAL SUMMARY (all sources)" -Level "INFO"
Write-Log "  Sources processed      : $($stats.SourcesProcessed) of $($stats.TotalSources)" -Level "INFO"
Write-Log "  Profiles found         : $($stats.TotalProfilesFound)" -Level "INFO"
if ($TestMode) {
    Write-Log "  Profiles to move/delete: $($stats.ProfilesToMoveCount) / $($stats.ProfilesToDeleteCount)" -Level "INFO"
    Write-Log "  FR to move/delete      : $($stats.FolderRedirectionToMoveCount) / $($stats.FolderRedirectionToDeleteCount)" -Level "INFO"
    Write-Log "  Orphaned FR reported   : $($stats.OrphanedFRReportedCount)" -Level "INFO"
    Write-Log "  Empty folders to delete: $($stats.EmptyFolders)" -Level "INFO"
}
else {
    $profMoveStr = Format-Count -Success $stats.ProfilesMovedSuccess -Total $stats.ProfilesToMoveCount
    $profDelStr = Format-Count -Success $stats.ProfilesDeletedSuccess -Total $stats.ProfilesToDeleteCount
    $frMoveStr = Format-Count -Success $stats.FolderRedirectionMovedSuccess -Total $stats.FolderRedirectionToMoveCount
    $frDelStr = Format-Count -Success $stats.FolderRedirectionDeletedSuccess -Total $stats.FolderRedirectionToDeleteCount
    Write-Log "  Profiles moved/deleted : $profMoveStr / $profDelStr" -Level "INFO"
    Write-Log "  FR moved/deleted       : $frMoveStr / $frDelStr" -Level "INFO"
    Write-Log "  Orphaned FR processed  : $($stats.OrphanedFRProcessedCount)" -Level "INFO"
    Write-Log "  Quarantine cleaned     : $($stats.OldQuarantineProfiles)" -Level "INFO"
    Write-Log "  Empty folders deleted  : $($stats.EmptyFolders)" -Level "INFO"
}
Write-Log "" -Level "INFO"

# Detailed per-source breakdown
foreach ($rep in $sourceReports) {
    Write-Log "--- $($rep.SourceName) ---" -Level "INFO"
    Write-Log "  Users processed        : $($rep.UsersProcessed)" -Level "INFO"
    Write-Log "  Profiles found         : $($rep.ProfilesFound)" -Level "INFO"
    Write-Log "    - Citrix UPM         : $($rep.CitrixProfiles)" -Level "INFO"
    Write-Log "    - MS Roaming         : $($rep.RoamingProfiles)" -Level "INFO"
    Write-Log "    - Reset Citrix       : $($rep.ResetProfiles)" -Level "INFO"
    Write-Log "    - Undefined/empty    : $($rep.NotDefinedProfiles)" -Level "INFO"

    if ($TestMode) {
        Write-Log "  Planned actions (TEST):" -Level "INFO"
        Write-Log "    Profiles: move $($rep.ProfilesToMove), delete $($rep.ProfilesToDelete)" -Level "INFO"
        Write-Log "    FR      : move $($rep.FRToMove), delete $($rep.FRToDelete)" -Level "INFO"
        Write-Log "    Orphaned FR reported : $($rep.OrphanedFRReported)" -Level "INFO"
    }
    else {
        $pMove = Format-Count -Success $rep.ProfilesMovedSuccess -Total $rep.ProfilesToMove
        $pDel = Format-Count -Success $rep.ProfilesDeletedSuccess -Total $rep.ProfilesToDelete
        $fMove = Format-Count -Success $rep.FRMovedSuccess -Total $rep.FRToMove
        $fDel = Format-Count -Success $rep.FRDeletedSuccess -Total $rep.FRToDelete
        Write-Log "  Executed actions:" -Level "INFO"
        Write-Log "    Profiles: moved $pMove (err $($rep.ProfilesMoveFailed)), deleted $pDel (err $($rep.ProfilesDeleteFailed))" -Level "INFO"
        Write-Log "    FR      : moved $fMove (err $($rep.FRMoveFailed)), deleted $fDel (err $($rep.FRDeleteFailed))" -Level "INFO"
        Write-Log "    Orphaned FR processed: $($rep.OrphanedFRProcessed)" -Level "INFO"
        Write-Log "    Quarantine cleaned   : $($rep.QuarantineCleaned)" -Level "INFO"
    }
    Write-Log "" -Level "INFO"
}

Write-Log "=" * 70 -Level "INFO"

# ========== SAVE REPORT ==========
$reportFile = Join-Path $LogDir "ProfileCleanup_Report_$(Get-Date -Format 'yyyyMMdd_HHmmss').txt"
$htmlReportFile = Join-Path $LogDir "ProfileCleanup_Report_$(Get-Date -Format 'yyyyMMdd_HHmmss').html"

# Text report (summary)
$reportContent = @"
Profile Cleanup Report
Execution Date: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')
Execution Mode: $(if($TestMode){'TESTING'}else{'REAL'})
==========================================================

TOTAL SUMMARY
-------------
Sources processed      : $($stats.SourcesProcessed) of $($stats.TotalSources)
Profiles found         : $($stats.TotalProfilesFound)
$(
if ($TestMode) {
@"
Profiles to move/delete: $($stats.ProfilesToMoveCount) / $($stats.ProfilesToDeleteCount)
FR to move/delete      : $($stats.FolderRedirectionToMoveCount) / $($stats.FolderRedirectionToDeleteCount)
Orphaned FR reported   : $($stats.OrphanedFRReportedCount)
Empty folders to delete: $($stats.EmptyFolders)
"@
} else {
@"
Profiles moved/deleted : $(Format-Count -Success $stats.ProfilesMovedSuccess -Total $stats.ProfilesToMoveCount) / $(Format-Count -Success $stats.ProfilesDeletedSuccess -Total $stats.ProfilesToDeleteCount)
FR moved/deleted       : $(Format-Count -Success $stats.FolderRedirectionMovedSuccess -Total $stats.FolderRedirectionToMoveCount) / $(Format-Count -Success $stats.FolderRedirectionDeletedSuccess -Total $stats.FolderRedirectionToDeleteCount)
Orphaned FR processed  : $($stats.OrphanedFRProcessedCount)
Quarantine cleaned     : $($stats.OldQuarantineProfiles)
Empty folders deleted  : $($stats.EmptyFolders)
"@
}
)

DETAILS BY SOURCE
-----------------
$(
foreach ($rep in $sourceReports) {
@"
[$($rep.SourceName)]
  Users processed        : $($rep.UsersProcessed)
  Profiles found         : $($rep.ProfilesFound)
    - Citrix UPM         : $($rep.CitrixProfiles)
    - MS Roaming         : $($rep.RoamingProfiles)
    - Reset Citrix       : $($rep.ResetProfiles)
    - Undefined/empty    : $($rep.NotDefinedProfiles)
$(
if ($TestMode) {
@"
  Planned (TEST):
    Profiles: move $($rep.ProfilesToMove), delete $($rep.ProfilesToDelete)
    FR      : move $($rep.FRToMove), delete $($rep.FRToDelete)
    Orphaned FR reported : $($rep.OrphanedFRReported)
"@
} else {
@"
  Executed:
    Profiles: moved $(Format-Count -Success $rep.ProfilesMovedSuccess -Total $rep.ProfilesToMove) (err $($rep.ProfilesMoveFailed)), deleted $(Format-Count -Success $rep.ProfilesDeletedSuccess -Total $rep.ProfilesToDelete) (err $($rep.ProfilesDeleteFailed))
    FR      : moved $(Format-Count -Success $rep.FRMovedSuccess -Total $rep.FRToMove) (err $($rep.FRMoveFailed)), deleted $(Format-Count -Success $rep.FRDeletedSuccess -Total $rep.FRToDelete) (err $($rep.FRDeleteFailed))
    Orphaned FR processed: $($rep.OrphanedFRProcessed)
    Quarantine cleaned   : $($rep.QuarantineCleaned)
"@
}
)

"@
}
)

SOURCES CONFIGURATION
---------------------
$($userProfileSources | Where-Object { $_.Enabled } | ForEach-Object {
    $qStatus = if ($_.EnableQuarantine -eq $false) { 'quarantine disabled' } else { 'quarantine enabled' }
    "- $($_.Name): $($_.ProfileRoot) (pattern: $($_.PatternProfile), $qStatus)"
} | Out-String)
==========================================================
Execution log: $LogFile
"@

$reportContent | Out-File -FilePath $reportFile -Encoding UTF8

# HTML report with detailed table (if there are entries)
if ($executionDetails.Count -gt 0) {
    $htmlRows = ""
    foreach ($item in $executionDetails) {
        $statusColor = switch ($item.Status) {
            "SUCCESS" { "#4CAF50" }  # green
            "FAILED" { "#f44336" }  # red
            "TEST" { "#ff9800" }  # orange
            "REPORTED" { "#2196F3" }  # blue (for reported only)
            default { "#ffffff" }
        }
        $htmlRows += @"
<tr style="background-color: $statusColor;">
    <td>$($item.Source)</td>
    <td>$($item.UserName)</td>
    <td>$($item.ProfileName)</td>
    <td>$($item.ProfileType)</td>
    <td>$($item.Reason)</td>
    <td>$($item.Details)</td>
    <td>$($item.Action)</td>
    <td>$($item.Destination)</td>
    <td>$($item.Status)</td>
    <td>$($item.ErrorMessage)</td>
</tr>
"@
    }

    $htmlTemplate = @"
<!DOCTYPE html>
<html>
<head><title>Profile Cleanup Report</title>
<style>
body { font-family: Arial, sans-serif; }
h2 { color: #333; }
table { border-collapse: collapse; width: 100%; margin-top: 20px; }
th, td { border: 1px solid #ddd; padding: 8px; text-align: left; }
th { background-color: #4CAF50; color: white; }
tr:nth-child(even) { background-color: #f9f9f9; }
</style>
</head>
<body>
<h2>Profile Cleanup $(if($TestMode){'TEST'}else{'EXECUTION'}) Report</h2>
<p><strong>Date:</strong> $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')</p>
<p><strong>Mode:</strong> $(if($TestMode){'TESTING'}else{'REAL'})</p>
<p><strong>Total items processed:</strong> $($executionDetails.Count)</p>
<table>
<tr>
    <th>Source</th>
    <th>User</th>
    <th>Profile</th>
    <th>Type</th>
    <th>Reason</th>
    <th>Details</th>
    <th>Action</th>
    <th>Destination</th>
    <th>Status</th>
    <th>Error</th>
</tr>
$htmlRows
</table>
</body>
</html>
"@
    $htmlTemplate | Out-File -FilePath $htmlReportFile -Encoding UTF8
    Write-Log "HTML report saved: $htmlReportFile" -Level "INFO"
}

Write-Log "Report saved: $reportFile" -Level "INFO"
Write-Log "Log saved: $LogFile" -Level "INFO"

if ($MailSettings.Enabled) {
    Send-ReportEmail -ReportFile $reportFile -HtmlReportFile $htmlReportFile -LogFile $LogFile -Stats $stats -TestMode $TestMode
}

if ($TestMode) {
    Write-Log "TESTING COMPLETED. Check the report before real execution." -Level "INFO"
}