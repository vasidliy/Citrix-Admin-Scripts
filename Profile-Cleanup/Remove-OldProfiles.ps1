#requires -version 5.1

<#
.SYNOPSIS
    Automated cleanup of old, corrupted, or orphaned user profiles and Folder Redirection data.
    Supports Citrix UPM, Microsoft Roaming, and reset profiles with configurable retention policies,
    quarantine, Active Directory validation, and detailed reporting.

.DESCRIPTION
    This script scans network shares containing user profiles and optionally redirected folders.
    It identifies profile types using configurable indicators, checks the health of NTUSER.DAT,
    evaluates age based on last write time (with optional fallback sources when NTUSER.DAT is missing),
    and queries Active Directory to determine if the user account still exists and is enabled.
    
    Based on fully customizable rules per source, profiles can be:
    - Moved to a quarantine folder (with automatic cleanup after a retention period).
    - Deleted immediately.
    - Ignored (left untouched).
    
    Separate actions can be defined for missing vs. undersized NTUSER.DAT files.
    
    Folder Redirection (FR) support:
    - Redirected folders associated with a cleaned‑up profile are also moved/deleted.
    - Orphaned redirected folders (no corresponding profile) can be reported, ignored, or cleaned up.
    
    Adaptive profile discovery:
    - The script can scan subfolders up to a configurable depth to locate profile indicators.
    - Indicators are configurable (e.g., "UPM_Profile\NTUSER.DAT", "NTUSER.DAT", "UPMSettings.ini").
    
    Comprehensive logging, console output, HTML reports, and email notifications are included.
    A test mode allows full simulation without modifying any files.

.PARAMETER ConfigFile
    Path to the JSON configuration file. If a relative path is given, it is resolved from the script's location.
    Default: "Remove-OldProfiles.ps1.json" in the script directory.

.EXAMPLE
    .\Remove-OldProfiles.ps1
    Runs the script using the default configuration file in the script's folder.

.EXAMPLE
    .\Remove-OldProfiles.ps1 -ConfigFile "C:\Configs\prod_cleanup.json"
    Runs the script with a custom configuration file located at the specified absolute path.

.EXAMPLE
    # Test mode enabled in configuration:
    # "GeneralSettings": { "TestMode": true }
    .\Remove-OldProfiles.ps1
    Executes a dry run – all actions are logged but no files are moved or deleted.

.NOTES
    Requirements:
    - PowerShell 5.1 or later.
    - Read/Write access to profile and quarantine shares (UNC paths supported).
    - ActiveDirectory PowerShell module (optional; required only if AD checks are enabled).
    - SMTP server accessible if email reporting is enabled.
    
    Long paths (>260 characters) and special characters (e.g., %3A) are fully supported via the \\?\UNC\ prefix.
    
    For detailed configuration instructions, refer to the accompanying README.md and example JSON file.
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
    
    # If detailed logging is disabled, skip DEBUG messages entirely
    if ($Level -eq "DEBUG" -and -not $GeneralSettings.DetailedLogging) {
        return
    }
    
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
                if ($GeneralSettings.DetailedLogging) {
                    Write-Host $logMessage.Trim() -ForegroundColor Gray 
                }
            }
        }
    }
}

# ========== EMAIL FUNCTION ==========
function Send-ReportEmail {
    param(
        [string]$HtmlReportFile, # attached HTML file (optional)
        [string]$HtmlBody, # HTML content for the email body
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
        
        $attachments = @()
        if ($HtmlReportFile -and (Test-Path $HtmlReportFile)) { 
            $attachments += $HtmlReportFile 
        }
        
        $mailParams = @{
            SmtpServer  = $MailSettings.SmtpServer
            Port        = if ($MailSettings.Port) { $MailSettings.Port } else { 25 }
            UseSsl      = if ($MailSettings.UseSSL) { $MailSettings.UseSSL } else { $false }
            From        = $MailSettings.From
            To          = $MailSettings.To
            Subject     = $subject
            Body        = $HtmlBody
            BodyAsHtml  = $true
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

# ========== PROFILE DISCOVERY FUNCTION (ADAPTIVE SCAN) ==========
function Find-UserProfiles {
    param(
        [string]$UserFolderPath,
        [int]$MaxDepth = 2,
        [string[]]$Indicators = @("UPM_Profile\NTUSER.DAT", "NTUSER.DAT", "UPMSettings.ini", "Pending", "AppData", "Desktop", "Documents"),
        [string[]]$UPMProfileMarkers = @("UPM_Profile\NTUSER.DAT", "UPMSettings.ini", "Pending"),
        [string[]]$MSRoamingMarkers = @("NTUSER.DAT", "AppData", "Desktop", "Documents")
    )

    # Log parameters once per source (first call)
    Write-Log "Using profile indicators: $($Indicators -join ', ')" -Level "DEBUG"
    Write-Log "Using UPM markers: $($UPMProfileMarkers -join ', ')" -Level "DEBUG"
    Write-Log "Using MS Roaming markers: $($MSRoamingMarkers -join ', ')" -Level "DEBUG"

    $userProfiles = [System.Collections.Generic.List[PSObject]]::new()

    function Test-AnyMarkerExists {
        param([string]$FolderPath, [string[]]$MarkerList)
        foreach ($marker in $MarkerList) {
            if (Test-Path -LiteralPath (Join-Path $FolderPath $marker)) {
                return $true
            }
        }
        return $false
    }

    function Test-IsProfileFolder {
        param(
            [string]$Path,
            [string[]]$Indicators,
            [string[]]$UPMProfileMarkers,
            [string[]]$MSRoamingMarkers
        )

        $folderName = Split-Path $Path -Leaf

        # 1. Reset Citrix profile (by name pattern)
        if ($folderName -match 'upm_' -and $folderName -notmatch 'UPM_Profile') {
            Write-Log "Detected reset Citrix profile by upm_ pattern in name: $Path" -Level "DEBUG"
            return @{
                IsProfile   = $true
                ProfileType = 'ResetCitrixProfile'
                NtuserPath  = ''
                Details     = 'Reset Citrix profile (by upm_ pattern in folder name)'
                StopScan    = $true
            }
        }

        # 2. Check all configured indicators – early exit on first match
        foreach ($indicator in $Indicators) {
            $indicatorPath = Join-Path $Path $indicator
            if (Test-Path -LiteralPath $indicatorPath) {
                # Determine type based on markers, NOT indicator name
                if (Test-AnyMarkerExists -FolderPath $Path -MarkerList $UPMProfileMarkers) {
                    $ntuser = Join-Path $Path "UPM_Profile\NTUSER.DAT"
                    $details = "Found Citrix UPM profile"
                    if (-not (Test-Path -LiteralPath $ntuser)) {
                        $details += " (NTUSER.DAT missing in UPM_Profile)"
                    }
                    Write-Log "Detected Citrix UPM profile via indicator '$indicator' in $Path" -Level "DEBUG"
                    return @{
                        IsProfile   = $true
                        ProfileType = 'CitrixUPM'
                        NtuserPath  = $ntuser
                        Details     = $details
                        StopScan    = $true
                    }
                }
                elseif (Test-AnyMarkerExists -FolderPath $Path -MarkerList $MSRoamingMarkers) {
                    $ntuserPath = Join-Path $Path "NTUSER.DAT"
                    $details = "Found MS Roaming profile"
                    if (-not (Test-Path -LiteralPath $ntuserPath)) {
                        $details += " (NTUSER.DAT missing but Roaming markers present)"
                        $ntuserPath = ''
                    }
                    Write-Log "Detected Microsoft Roaming profile via indicator '$indicator' in $Path" -Level "DEBUG"
                    return @{
                        IsProfile   = $true
                        ProfileType = 'MicrosoftRoaming'
                        NtuserPath  = $ntuserPath
                        Details     = $details
                        StopScan    = $true
                    }
                }
                else {
                    # Indicator found, but cannot classify -> NotDefined
                    Write-Log "Indicator '$indicator' found in $Path but profile type cannot be determined, treating as NotDefined" -Level "DEBUG"
                    return @{
                        IsProfile   = $true
                        ProfileType = 'NotDefined'
                        NtuserPath  = ''
                        Details     = "Profile indicator '$indicator' found but type cannot be determined (no UPM or MS Roaming markers)"
                        StopScan    = $true
                    }
                }
            }
        }

        return @{ IsProfile = $false; StopScan = $false }
    }

    function Search-ProfileFolder {
        param(
            [string]$CurrentPath,
            [int]$CurrentDepth,
            [string[]]$Indicators,
            [string[]]$UPMProfileMarkers,
            [string[]]$MSRoamingMarkers
        )

        if ($CurrentDepth -gt $MaxDepth) { return }

        Write-Log "Scanning depth $CurrentDepth : $CurrentPath" -Level "DEBUG"

        $result = Test-IsProfileFolder -Path $CurrentPath -Indicators $Indicators -UPMProfileMarkers $UPMProfileMarkers -MSRoamingMarkers $MSRoamingMarkers
        if ($result.IsProfile) {
            $userProfiles.Add([PSCustomObject]@{
                    Path        = $CurrentPath
                    ProfileType = $result.ProfileType
                    NtuserPath  = $result.NtuserPath
                    Details     = $result.Details
                })
            if ($result.StopScan) {
                Write-Log "Stop scanning subfolders of $CurrentPath (profile already identified)" -Level "DEBUG"
                return
            }
        }

        $subDirs = Get-ChildItem -LiteralPath $CurrentPath -Directory -Force -ErrorAction SilentlyContinue
        foreach ($subDir in $subDirs) {
            Search-ProfileFolder -CurrentPath $subDir.FullName -CurrentDepth ($CurrentDepth + 1) -Indicators $Indicators -UPMProfileMarkers $UPMProfileMarkers -MSRoamingMarkers $MSRoamingMarkers
        }
    }

    Search-ProfileFolder -CurrentPath $UserFolderPath -CurrentDepth 0 -Indicators $Indicators -UPMProfileMarkers $UPMProfileMarkers -MSRoamingMarkers $MSRoamingMarkers

    if ($userProfiles.Count -eq 0) {
        Write-Log "No profile indicators found in $UserFolderPath (max depth $MaxDepth)" -Level "DEBUG"
        $userProfiles.Add([PSCustomObject]@{
                Path        = $UserFolderPath
                ProfileType = 'NotDefined'
                NtuserPath  = ''
                Details     = "No profile indicators found within depth $MaxDepth"
            })
    }

    return $userProfiles
}

# ========== PROFILE HEALTH CHECK FUNCTION ==========
function Test-ProfileHealth {
    param(
        [string]$NtuserPath,
        [int]$MinNtuserSizeKB = 1,
        [string]$FallbackPath = $null
    )
    
    Write-Log "Checking health for path: $NtuserPath" -Level "DEBUG"
    
    $result = @{
        IsHealthy         = $false
        Reason            = ""
        ReasonCode        = ""
        SizeKB            = 0
        LastWriteTime     = $null
        FallbackUsed      = $false
        FallbackLastWrite = $null
    }

    # Helper to safely get LastWriteTime from a fallback path
    function Get-FallbackTimestamp {
        param([string]$Path)
        if ($Path -and (Test-Path -LiteralPath $Path)) {
            try {
                $item = Get-Item -LiteralPath $Path -Force -ErrorAction Stop
                return $item.LastWriteTime
            }
            catch {
                Write-Log "Could not access fallback path '$Path': $_" -Level "DEBUG"
            }
        }
        return $null
    }

    # If NtuserPath is empty or doesn't exist, immediately try fallback
    if ([string]::IsNullOrEmpty($NtuserPath)) {
        Write-Log "NTUSER.DAT path is empty" -Level "DEBUG"
        $result.Reason = "NTUSER.DAT not specified"
        $result.ReasonCode = "Missing"
        $result.FallbackLastWrite = Get-FallbackTimestamp -Path $FallbackPath
        if ($result.FallbackLastWrite) {
            Write-Log "Fallback timestamp obtained from '$FallbackPath': $($result.FallbackLastWrite)" -Level "DEBUG"
        }
        return $result
    }

    if (-not (Test-Path -LiteralPath $NtuserPath)) {
        Write-Log "NTUSER.DAT does not exist at path: $NtuserPath" -Level "DEBUG"
        $result.Reason = "NTUSER.DAT not found"
        $result.ReasonCode = "Missing"
        $result.FallbackLastWrite = Get-FallbackTimestamp -Path $FallbackPath
        if ($result.FallbackLastWrite) {
            Write-Log "Fallback timestamp obtained from '$FallbackPath': $($result.FallbackLastWrite)" -Level "DEBUG"
        }
        return $result
    }

    # File exists – try to read properties
    try {
        $fileInfo = [System.IO.FileInfo]::new($NtuserPath)
        $fileInfo.Refresh()
        if (-not $fileInfo.Exists) {
            throw "File does not exist (after Refresh)"
        }
        $sizeKB = [math]::Round($fileInfo.Length / 1KB, 2)
        
        if ($MinNtuserSizeKB -gt 0 -and $sizeKB -lt $MinNtuserSizeKB) {
            $result.Reason = "NTUSER.DAT too small ($sizeKB KB < $MinNtuserSizeKB KB)"
            $result.ReasonCode = "TooSmall"
            $result.SizeKB = $sizeKB
            $result.LastWriteTime = $fileInfo.LastWriteTime
            $result.FallbackLastWrite = Get-FallbackTimestamp -Path $FallbackPath
            if ($result.FallbackLastWrite) {
                Write-Log "Fallback timestamp obtained from '$FallbackPath': $($result.FallbackLastWrite)" -Level "DEBUG"
            }
            return $result
        }
        
        # Healthy
        $result.IsHealthy = $true
        $result.Reason = "OK"
        $result.ReasonCode = "Healthy"
        $result.SizeKB = $sizeKB
        $result.LastWriteTime = $fileInfo.LastWriteTime
        return $result
    }
    catch {
        Write-Log "Error accessing NTUSER.DAT: $_" -Level "ERROR"
        $result.Reason = "Error accessing NTUSER.DAT: $_"
        $result.ReasonCode = "Error"
        $result.FallbackLastWrite = Get-FallbackTimestamp -Path $FallbackPath
        if ($result.FallbackLastWrite) {
            Write-Log "Fallback timestamp obtained from '$FallbackPath': $($result.FallbackLastWrite)" -Level "DEBUG"
        }
        return $result
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
                (Get-Item -LiteralPath $DestinationPath -Force).LastWriteTime = Get-Date
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
            "/E", # copy subdirectories, including empty ones
            "/MOV", # move files (delete from source after copying)
            "/R:3", # retry 3 times on failure
            "/W:5", # wait 5 seconds between retries
            "/NP", # no progress
            "/NJH", # no job header
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
                (Get-Item -LiteralPath $DestinationPath -Force).LastWriteTime = Get-Date
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
        [string]$SourceName
    )
    
    $cleanedCount = 0
    
    if (-not (Test-Path $QuarantinePath)) {
        Write-Log "Quarantine folder does not exist: $QuarantinePath" -Level "DEBUG"
        return 0
    }
    
    $cutoffDate = (Get-Date).AddDays(-$QuarantineDays)
    $oldProfiles = Get-ChildItem -Path $QuarantinePath -Directory -Force -ErrorAction SilentlyContinue |
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
                        $longPath = '\\?\UNC\' + $profileFullPath.Substring(2)
                    }
                    else {
                        $longPath = '\\?\' + $profileFullPath
                    }
                    Remove-Item -LiteralPath $longPath -Recurse -Force -ErrorAction Stop
                    Write-Log "Deleted old profile from quarantine: $profileName" -Level "INFO"
                    $status = "SUCCESS"
                }
                
                $detail = [PSCustomObject]@{
                    Source       = $SourceName
                    UserName     = ""
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
                $cleanedCount++
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
                }
            }
        }
    }
    else {
        Write-Log "No old profiles found in quarantine" -Level "DEBUG"
    }
    
    return $cleanedCount
}

# ========== ORPHANED FOLDER REDIRECTION CLEANUP FUNCTION ==========
function Clear-OrphanedFolderRedirection {
    param(
        [string[]]$FolderRedirectionPaths,
        [string]$QuarantinePath,
        [bool]$EnableQuarantine,
        [string]$SourceName,
        [hashtable]$FoundUsers, # usernames that have a corresponding profile folder
        [string]$ProcessOrphanedFR, # "Disabled", "ReportOnly", or "Delete"
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

        $frSubFolders = Get-ChildItem -Path $frPath -Directory -Force -ErrorAction SilentlyContinue
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

                    if ($TestMode) {
                        $frDetail.Status = "TEST"
                        Write-Log "TEST: Orphaned FR would be moved: $($frSubFolder.FullName) -> $destinationPath" -Level "INFO"
                        $script:stats.OrphanedFRReportedCount++
                        $script:stats.FolderRedirectionToMoveCount++
                    }
                    else {
                        $moved = Move-FolderRedirectionToQuarantine -SourcePath $frSubFolder.FullName -DestinationPath $destinationPath -UserName $userName -Reason "ORPHANED"
                        if ($moved) {
                            $frDetail.Status = "SUCCESS"
                            $script:stats.FolderRedirectionMovedSuccess++
                            $script:stats.OrphanedFRProcessedCount++
                            $script:stats.FolderRedirectionToMoveCount++
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

                    if ($TestMode) {
                        $frDetail.Status = "TEST"
                        Write-Log "TEST: Orphaned FR would be deleted: $($frSubFolder.FullName)" -Level "INFO"
                        $script:stats.OrphanedFRReportedCount++
                        $script:stats.FolderRedirectionToDeleteCount++
                    }
                    else {
                        $removed = Remove-FolderRedirection -UserFolderPath $frSubFolder.FullName -UserName $userName -Reason "ORPHANED"
                        if ($removed) {
                            $frDetail.Status = "SUCCESS"
                            $script:stats.FolderRedirectionDeletedSuccess++
                            $script:stats.OrphanedFRProcessedCount++
                            $script:stats.FolderRedirectionToDeleteCount++
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
        [string[]]$ExcludePatterns, # patterns for excluded folder names
        [string]$SourceName, # added to link action to source in report
        [bool]$SkipNewFolders = $false, # if true, skip folders created recently (unless aggressive actions)
        [int]$DaysToDelete = 365, # inactivity threshold used for SkipNewFolders
        [bool]$IsAggressive = $false        # if true, even new folders will be processed
    )
    
    $deletedCount = 0
    
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

    # Apply SkipNewFolders filter (only if aggressive actions are not set)
    if ($SkipNewFolders -and -not $IsAggressive) {
        $cutoffDate = (Get-Date).AddDays(-$DaysToDelete)
        $subFolders = $subFolders | Where-Object { $_.CreationTime -le $cutoffDate }
        Write-Log "Empty folder cleanup: SkipNewUserFolders enabled (non-aggressive), ignoring folders created after $cutoffDate" -Level "INFO"
    }
    
    foreach ($subFolder in $subFolders) {
        # Check if the folder is empty (including hidden items)
        $items = Get-ChildItem -Path $subFolder.FullName -Force -ErrorAction SilentlyContinue
        if (-not $items) {
            $status = $null
            $errorMsg = $null
            if ($TestMode) {
                Write-Log "TEST: Empty folder would be deleted: $($subFolder.FullName)" -Level "INFO"
                $status = "TEST"
            }
            else {
                try {
                    Remove-Item -Path $subFolder.FullName -Force -ErrorAction SilentlyContinue
                    Write-Log "Deleted empty folder: $($subFolder.FullName)" -Level "INFO"
                    $status = "SUCCESS"
                }
                catch {
                    Write-Log "Could not delete empty folder $($subFolder.FullName): $_" -Level "WARN"
                    $status = "FAILED"
                    $errorMsg = $_.Exception.Message
                }
            }
            
            # Add to execution details for HTML report
            $detail = [PSCustomObject]@{
                Source       = $SourceName
                UserName     = ""
                ProfileName  = $subFolder.Name
                ProfileType  = "EmptyFolder"
                Reason       = "EMPTY_FOLDER"
                Details      = "Empty user folder"
                Action       = "DELETE"
                Destination  = "N/A"
                Status       = $status
                ErrorMessage = $errorMsg
            }
            $executionDetails.Add($detail)
            
            $deletedCount++
        }
    }
    
    return $deletedCount
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

    # Actions for corrupted profiles (Missing and TooSmall)
    $actionMissing = "Quarantine"  # default
    if ($source.PSObject.Properties.Name -contains 'ActionOnMissingNTUserDat') {
        $actionMissing = $source.ActionOnMissingNTUserDat
    }
    $actionTooSmall = "Quarantine"  # default
    if ($source.PSObject.Properties.Name -contains 'ActionOnTooSmallNTUserDat') {
        $actionTooSmall = $source.ActionOnTooSmallNTUserDat
    }
    Write-Log "Action on missing NTUSER.DAT: $actionMissing" -Level "INFO"
    Write-Log "Action on too small NTUSER.DAT: $actionTooSmall" -Level "INFO"

    # Fallback age settings
    $useFallbackAge = $true   # default
    if ($source.PSObject.Properties.Name -contains 'UseFallbackAgeWhenNTUserMissing') {
        $useFallbackAge = [bool]$source.UseFallbackAgeWhenNTUserMissing
    }
    $fallbackSource = "Auto"  # default
    if ($source.PSObject.Properties.Name -contains 'FallbackAgeSource') {
        $fallbackSource = $source.FallbackAgeSource
    }
    Write-Log "Use fallback age when NTUSER.DAT missing: $useFallbackAge" -Level "INFO"
    Write-Log "Fallback age source: $fallbackSource" -Level "INFO"

    # Adaptive scan settings
    $maxScanDepth = 2   # default
    if ($source.PSObject.Properties.Name -contains 'MaxProfileScanDepth') {
        $maxScanDepth = [int]$source.MaxProfileScanDepth
    }

    # Read optional parameters from JSON; if absent, $null is passed, function defaults are used
    $profileIndicators = if ($source.PSObject.Properties.Name -contains 'ProfileIndicators') {
        $source.ProfileIndicators
    }
    else { $null }
    $upmMarkers = if ($source.PSObject.Properties.Name -contains 'UPMProfileMarkers') {
        $source.UPMProfileMarkers
    }
    else { $null }
    $msRoamingMarkers = if ($source.PSObject.Properties.Name -contains 'MSRoamingMarkers') {
        $source.MSRoamingMarkers
    }
    else { $null }

    Write-Log "Max profile scan depth: $maxScanDepth" -Level "INFO"

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
        EmptyFoldersCleaned    = 0 
    }

    try {
        if (-not (Test-Path $source.ProfileRoot)) {
            Write-Log "ERROR: Network path is not accessible" -Level "ERROR"
            continue
        }
        
        $userFolders = Get-ChildItem -Path $source.ProfileRoot -Directory -Force -ErrorAction SilentlyContinue |
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
            # Build a single regex from wildcard patterns
            $regexPatterns = $source.ExcludeFolders | ForEach-Object {
                # Escape all regex special characters, then convert wildcards
                $escaped = [regex]::Escape($_)
                # Replace escaped wildcards: \* -> .* and \? -> .
                $escaped = $escaped -replace '\\\*', '.*' -replace '\\\?', '.'
                # Anchor to full folder name
                "^$escaped$"
            }
            $excludeRegex = $regexPatterns -join '|'
            $userFolders = $userFolders | Where-Object { $_.Name -notmatch $excludeRegex }
            Write-Log "After exclusions, remaining folders: $($userFolders.Count)" -Level "INFO"
        }

        # NEW: Store all candidate folders before SkipNewUserFolders filter
        $allCandidateFolders = $userFolders

        # NEW: Build $foundUsers from ALL candidate folders (correct orphaned FR detection)
        $foundUsers = @{}
        foreach ($folder in $allCandidateFolders) {
            $rawName = $folder.Name
            $userName = $rawName -replace '\.v\d+$', '' -replace '^upm_', '' -replace "^$sourceNetBios\.", '' -replace "^$sourceFQDN\.", '' -replace "\.$sourceNetBios$", '' -replace "\.$sourceFQDN$", ''
            $foundUsers[$userName] = $true
        }

        if (-not $allCandidateFolders) {
            Write-Log "No user folders found (or all excluded)" -Level "INFO"
            # Even if no user folders, we can still check for orphaned FR
            if ($orphanFRMode -ne "Disabled" -and $source.PSObject.Properties.Name -contains 'FolderRedirectionPaths') {
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

        Write-Log "Candidate user folders found: $($allCandidateFolders.Count)" -Level "INFO"

        # NEW: Apply SkipNewUserFolders filter (one-time instead of per-folder)
        $skipNewFolders = $false
        if ($source.PSObject.Properties.Name -contains 'SkipNewUserFolders') {
            $skipNewFolders = [bool]$source.SkipNewUserFolders
        }

        $foldersToProcess = $allCandidateFolders
        if ($skipNewFolders) {
            $aggressiveActions = @('Quarantine', 'Delete')
            $isAggressive = ($actionMissing -in $aggressiveActions) -or ($actionTooSmall -in $aggressiveActions)
            if (-not $isAggressive) {
                $cutoffDate = (Get-Date).AddDays(-$source.DaysToDelete)
                $foldersToProcess = $foldersToProcess | Where-Object { $_.CreationTime -le $cutoffDate }
                Write-Log "SkipNewUserFolders (non-aggressive): filtered out folders created after $cutoffDate. Remaining: $($foldersToProcess.Count)" -Level "INFO"
            }
            else {
                Write-Log "SkipNewUserFolders enabled but aggressive actions are set - processing all candidate folders" -Level "INFO"
            }
        }
        else {
            Write-Log "SkipNewUserFolders is disabled" -Level "DEBUG"
        }

        # NEW: Sort folders by name for predictable processing order
        $foldersToProcess = $foldersToProcess | Sort-Object -Property Name
        Write-Log "Folders to process: $($foldersToProcess.Count)" -Level "INFO"
        # END NEW

        foreach ($userFolder in $foldersToProcess) {
            # Extract username
            $rawName = $userFolder.Name
            $userName = $rawName `
                -replace '\.v\d+$', '' `
                -replace '^upm_', '' `
                -replace "^$sourceNetBios\.", '' `
                -replace "^$sourceFQDN\.", '' `
                -replace "\.$sourceNetBios$", '' `
                -replace "\.$sourceFQDN$", ''

            # NOTE: $foundUsers is already populated from all candidate folders above

            $sourceStats.UsersProcessed++

            if ($GeneralSettings.DetailedLogging -or $TestMode) {
                Write-Log "  User: $userName (original folder: $rawName)" -Level "DEBUG"
            }
            
            # Build parameter splat – only pass what was explicitly configured
            $findParams = @{
                UserFolderPath = $userFolder.FullName
                MaxDepth       = $maxScanDepth
            }
            if ($profileIndicators) { $findParams.Indicators = $profileIndicators }
            if ($upmMarkers) { $findParams.UPMProfileMarkers = $upmMarkers }
            if ($msRoamingMarkers) { $findParams.MSRoamingMarkers = $msRoamingMarkers }

            $userProfiles = Find-UserProfiles @findParams
           
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
                    # Determine fallback path (always set when possible, even if $useFallbackAge is false)
                    $profileRootFolder = $userProfile.Path
                    $fallbackPath = $null
                    
                    if ($userProfileType -eq "MicrosoftRoaming") {
                        # For MS Roaming, we can always fall back to the profile folder itself
                        $fallbackPath = $profileRootFolder
                    }
                    elseif ($userProfileType -eq "CitrixUPM") {
                        # For Citrix UPM, try UPMSettings.ini or the folder itself based on settings
                        $upmSettingsPath = Join-Path $profileRootFolder "UPMSettings.ini"
                        if ($useFallbackAge) {
                            if ($fallbackSource -eq "Auto" -or $fallbackSource -eq "UPMSettings") {
                                if (Test-Path -LiteralPath $upmSettingsPath) {
                                    $fallbackPath = $upmSettingsPath
                                }
                            }
                            if (-not $fallbackPath -and ($fallbackSource -eq "Auto" -or $fallbackSource -eq "Folder")) {
                                $fallbackPath = $profileRootFolder
                            }
                        }
                        else {
                            # Fallback age is disabled, but we still want fallback for health check completeness
                            $fallbackPath = $profileRootFolder
                        }
                    }
                    # For ResetCitrixProfile and NotDefined, $fallbackPath remains $null (no fallback needed)

                    # Check profile health
                    $health = Test-ProfileHealth -NtuserPath $userProfile.NtuserPath -MinNtuserSizeKB 1 -FallbackPath $fallbackPath

                    if (-not $health.IsHealthy) {
                        # Profile is corrupted (Missing or TooSmall)
                        $reasonCode = $health.ReasonCode
                        $actionPreference = if ($reasonCode -eq "Missing") { $actionMissing } else { $actionTooSmall }

                        # If fallback age is enabled and we have a valid fallback timestamp, check age first
                        if ($useFallbackAge -and $health.FallbackLastWrite) {
                            $lastWrite = $health.FallbackLastWrite
                            $cutoffDate = (Get-Date).AddDays(-$source.DaysToDelete)
                            $daysOld = [math]::Round(((Get-Date) - $lastWrite).TotalDays, 1)

                            if ($lastWrite -lt $cutoffDate) {
                                # Profile is old based on fallback source
                                $moveReason = "PROFILE_OLD"
                                $details = "Inactive $daysOld days (threshold: $($source.DaysToDelete) days, source: fallback)"
                                # Override action preference – use standard quarantine/delete for old profiles
                                $actionPreference = if ($enableQuarantine) { "Quarantine" } else { "Delete" }
                            }
                            else {
                                # Not old – apply the original Missing/TooSmall preference
                                if ($actionPreference -eq "Ignore") {
                                    Write-Log "Profile [$userProfileType] is corrupted ($($health.Reason)) but not old and action is Ignore, skipping: $($userProfile.Path)" -Level "INFO"
                                    continue
                                }
                                $moveReason = "PROFILE_CORRUPTED"
                                $details = $health.Reason
                            }
                        }
                        else {
                            # No fallback available or fallback disabled – use Missing/TooSmall action directly
                            if ($actionPreference -eq "Ignore") {
                                Write-Log "Profile [$userProfileType] is corrupted ($($health.Reason)) and action is Ignore, skipping: $($userProfile.Path)" -Level "INFO"
                                continue
                            }
                            $moveReason = "PROFILE_CORRUPTED"
                            $details = $health.Reason
                        }
                    }
                    else {
                        # Profile is healthy – standard age check
                        $lastWrite = $health.LastWriteTime
                        $cutoffDate = (Get-Date).AddDays(-$source.DaysToDelete)
                        $daysOld = [math]::Round(((Get-Date) - $lastWrite).TotalDays, 1)

                        if ($lastWrite -lt $cutoffDate) {
                            $moveReason = "PROFILE_OLD"
                            $details = "Inactive $daysOld days (threshold: $($source.DaysToDelete) days, source: NTUSER.DAT)"
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
                            else {
                                # Profile is current (active user, not old, not corrupted)
                                if ($GeneralSettings.DetailedLogging -or $TestMode) {
                                    $lastWriteSource = if ($health.FallbackUsed -and $health.FallbackLastWrite) { "fallback ($fallbackSource)" } else { "NTUSER.DAT" }
                                    $lastWriteDate = if ($health.FallbackUsed) { $health.FallbackLastWrite } else { $health.LastWriteTime }
                                    Write-Log "  Profile is current: $($userProfile.Path) (last write: $lastWriteDate [$lastWriteSource])" -Level "DEBUG"
                                }
                            }
                        }
                    }
                }
                
                # If there is a reason – profile is subject to processing
                if ($moveReason) {
                    $userHadAction = $true
                    
                    # Determine action: MOVE (quarantine) or DELETE
                    $action = if ($enableQuarantine) { "MOVE" } else { "DELETE" }
                    
                    # Override for corrupted profiles based on action preference (if not already overridden by age)
                    if ($moveReason -eq "PROFILE_CORRUPTED") {
                        $pref = if ($health.ReasonCode -eq "Missing") { $actionMissing } else { $actionTooSmall }
                        if ($pref -eq "Delete") {
                            $action = "DELETE"
                        }
                        elseif ($pref -eq "Quarantine" -and $enableQuarantine) {
                            $action = "MOVE"
                        }
                        elseif ($pref -eq "Quarantine" -and -not $enableQuarantine) {
                            $action = "DELETE"
                        }
                    }
                    
                    # For MOVE, build destination path
                    $destinationPath = $null
                    if ($action -eq "MOVE") {
                        $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
                        $safeUserName = $userName -replace '[\\/:*?"<>|]', '_'
                        $safeProfileName = $userProfileName -replace '[\\/:*?"<>|]', '_'
    
                        # Truncate username if too long (keep first 20 chars)
                        if ($safeUserName.Length -gt 20) {
                            $safeUserName = $safeUserName.Substring(0, 20)
                        }
                        # Truncate profile name if too long (keep first 20 chars)
                        if ($safeProfileName.Length -gt 20) {
                            $safeProfileName = $safeProfileName.Substring(0, 20)
                        }
    
                        if ($userProfileType -eq "ResetCitrixProfile") {
                            $quarantineFolder = "${safeUserName}_RESET_${safeProfileName}_${timestamp}"
                        }
                        else {
                            $quarantineFolder = "${safeUserName}_${safeProfileName}_${userProfileType}_${timestamp}_${moveReason}"
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
            $qCleaned = Clear-OldQuarantine -QuarantinePath $source.QuarantinePath -QuarantineDays $source.QuarantineDays -SourceName $source.Name
            $stats.OldQuarantineProfiles += $qCleaned
            $sourceStats.QuarantineCleaned += $qCleaned
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

            # Determine if aggressive actions are set (for SkipNewFolders logic)
            $aggressiveActions = @('Quarantine', 'Delete')
            $isAggressive = ($actionMissing -in $aggressiveActions) -or ($actionTooSmall -in $aggressiveActions)

            $emptyDeleted = Remove-EmptyFolders -FolderPath $source.ProfileRoot `
                -ExcludePatterns $excludePatterns `
                -SourceName $source.Name `
                -SkipNewFolders:$skipNewFolders `
                -DaysToDelete $source.DaysToDelete `
                -IsAggressive:$isAggressive

            $stats.EmptyFolders += $emptyDeleted
            $sourceStats.EmptyFoldersCleaned += $emptyDeleted
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
            EmptyFoldersCleaned    = $sourceStats.EmptyFoldersCleaned
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
        Write-Log "    Empty folders cleaned: $($rep.EmptyFoldersCleaned)" -Level "INFO"
    }
    Write-Log "" -Level "INFO"
}

Write-Log "=" * 70 -Level "INFO"

# ========== SAVE REPORT ==========
# Summary HTML report file
$htmlReportFile = Join-Path $LogDir "ProfileCleanup_Report_$(Get-Date -Format 'yyyyMMdd_HHmmss').html"

# Summary HTML report file (same as email body)
$summaryHtmlFile = Join-Path $LogDir "ProfileCleanup_Summary_$(Get-Date -Format 'yyyyMMdd_HHmmss').html"

# ========== PREPARE HTML EMAIL BODY ==========
$htmlBody = @"
<!DOCTYPE html>
<html>
<head>
<style>
body { font-family: Arial, sans-serif; font-size: 12px; margin: 10px; }
h2 { color: #333; font-size: 16px; margin: 0 0 5px 0; }
h3 { color: #555; font-size: 14px; margin: 10px 0 5px 0; }
table { border-collapse: collapse; width: auto; margin-bottom: 10px; }
th, td { border: 1px solid #aaa; padding: 3px 5px; text-align: left; vertical-align: top; }
th { background-color: #4CAF50; color: white; font-weight: bold; }
th:first-child, td:first-child { width: 210px; }
tr.section-header td { background-color: #e0e0e0; font-weight: bold; padding: 4px 5px; }
.value-success { color: #2e7d32; font-weight: bold; }
.value-error { color: #c62828; font-weight: bold; }
.value-test { color: #e65100; font-weight: bold; }
.value-neutral { color: #1565c0; font-weight: bold; }
.indent { padding-left: 15px !important; }
</style>
</head>
<body>
<h2>Profile Cleanup $(if($TestMode){'TEST'}else{'EXECUTION'}) Report</h2>
<p style="margin:0 0 5px 0;"><strong>Date:</strong> $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') &nbsp;|&nbsp; <strong>Mode:</strong> $(if($TestMode){'TESTING'}else{'REAL'})</p>

<h3>TOTAL SUMMARY</h3>
<table>
<tr><th>Metric</th><th>Value</th></tr>
<tr><td>Sources processed</td><td>$($stats.SourcesProcessed) of $($stats.TotalSources)</td></tr>
<tr><td>Profiles found</td><td>$($stats.TotalProfilesFound)</td></tr>
$(
if ($TestMode) {
@"
<tr><td>Profiles to move / delete</td><td><span class="value-test">$($stats.ProfilesToMoveCount) / $($stats.ProfilesToDeleteCount)</span></td></tr>
<tr><td>FR to move / delete</td><td><span class="value-test">$($stats.FolderRedirectionToMoveCount) / $($stats.FolderRedirectionToDeleteCount)</span></td></tr>
<tr><td>Orphaned FR reported</td><td><span class="value-test">$($stats.OrphanedFRReportedCount)</span></td></tr>
<tr><td>Quarantine to clean</td><td><span class="value-test">$($stats.OldQuarantineProfiles)</span></td></tr>
<tr><td>Empty folders to delete</td><td><span class="value-test">$($stats.EmptyFolders)</span></td></tr>
"@
} else {
    $profMoveStr = Format-Count -Success $stats.ProfilesMovedSuccess -Total $stats.ProfilesToMoveCount
    $profDelStr  = Format-Count -Success $stats.ProfilesDeletedSuccess -Total $stats.ProfilesToDeleteCount
    $frMoveStr   = Format-Count -Success $stats.FolderRedirectionMovedSuccess -Total $stats.FolderRedirectionToMoveCount
    $frDelStr    = Format-Count -Success $stats.FolderRedirectionDeletedSuccess -Total $stats.FolderRedirectionToDeleteCount
@"
<tr><td>Profiles moved / deleted</td><td><span class="value-success">$profMoveStr</span> / <span class="value-success">$profDelStr</span></td></tr>
<tr><td>FR moved / deleted</td><td><span class="value-success">$frMoveStr</span> / <span class="value-success">$frDelStr</span></td></tr>
<tr><td>Orphaned FR processed</td><td><span class="value-success">$($stats.OrphanedFRProcessedCount)</span></td></tr>
<tr><td>Quarantine cleaned</td><td><span class="value-neutral">$($stats.OldQuarantineProfiles)</span></td></tr>
<tr><td>Empty folders deleted</td><td><span class="value-neutral">$($stats.EmptyFolders)</span></td></tr>
"@
}
)
</table>

<h3>DETAILS BY SOURCE</h3>
<table>
<tr><th>Metric</th><th>Value</th></tr>
$(
foreach ($rep in $sourceReports) {
    # Source header
@"
<tr class="section-header"><td colspan="2">$($rep.SourceName)</td></tr>
<tr><td class="indent">Users processed</td><td>$($rep.UsersProcessed)</td></tr>
<tr><td class="indent">Profiles found</td><td>$($rep.ProfilesFound)</td></tr>
<tr><td style="padding-left:25px !important;">- Citrix UPM</td><td>$($rep.CitrixProfiles)</td></tr>
<tr><td style="padding-left:25px !important;">- MS Roaming</td><td>$($rep.RoamingProfiles)</td></tr>
<tr><td style="padding-left:25px !important;">- Reset Citrix</td><td>$($rep.ResetProfiles)</td></tr>
<tr><td style="padding-left:25px !important;">- Undefined/empty</td><td>$($rep.NotDefinedProfiles)</td></tr>
"@
    if ($TestMode) {
@"
<tr><td class="indent"><strong>Planned (TEST):</strong></td><td></td></tr>
<tr><td style="padding-left:30px !important;">Profiles: move</td><td><span class="value-test">$($rep.ProfilesToMove)</span></td></tr>
<tr><td style="padding-left:30px !important;">Profiles: delete</td><td><span class="value-test">$($rep.ProfilesToDelete)</span></td></tr>
<tr><td style="padding-left:30px !important;">FR: move</td><td><span class="value-test">$($rep.FRToMove)</span></td></tr>
<tr><td style="padding-left:30px !important;">FR: delete</td><td><span class="value-test">$($rep.FRToDelete)</span></td></tr>
<tr><td style="padding-left:30px !important;">Orphaned FR reported</td><td><span class="value-test">$($rep.OrphanedFRReported)</span></td></tr>
<tr><td style="padding-left:30px !important;">Quarantine to clean</td><td><span class="value-test">$($rep.QuarantineCleaned)</span></td></tr>
<tr><td style="padding-left:30px !important;">Empty folders to delete</td><td><span class="value-test">$($rep.EmptyFoldersCleaned)</span></td></tr>
"@
    } else {
        $pMove = Format-Count -Success $rep.ProfilesMovedSuccess -Total $rep.ProfilesToMove
        $pDel  = Format-Count -Success $rep.ProfilesDeletedSuccess -Total $rep.ProfilesToDelete
        $fMove = Format-Count -Success $rep.FRMovedSuccess -Total $rep.FRToMove
        $fDel  = Format-Count -Success $rep.FRDeletedSuccess -Total $rep.FRToDelete
@"
<tr><td class="indent"><strong>Executed:</strong></td><td></td></tr>
<tr><td style="padding-left:30px !important;">Profiles: moved</td><td><span class="value-success">$pMove</span>$(if($rep.ProfilesMoveFailed -gt 0) {" <span class='value-error'>(err $($rep.ProfilesMoveFailed))</span>"})</td></tr>
<tr><td style="padding-left:30px !important;">Profiles: deleted</td><td><span class="value-success">$pDel</span>$(if($rep.ProfilesDeleteFailed -gt 0) {" <span class='value-error'>(err $($rep.ProfilesDeleteFailed))</span>"})</td></tr>
<tr><td style="padding-left:30px !important;">FR: moved</td><td><span class="value-success">$fMove</span>$(if($rep.FRMoveFailed -gt 0) {" <span class='value-error'>(err $($rep.FRMoveFailed))</span>"})</td></tr>
<tr><td style="padding-left:30px !important;">FR: deleted</td><td><span class="value-success">$fDel</span>$(if($rep.FRDeleteFailed -gt 0) {" <span class='value-error'>(err $($rep.FRDeleteFailed))</span>"})</td></tr>
<tr><td style="padding-left:30px !important;">Orphaned FR processed</td><td><span class="value-success">$($rep.OrphanedFRProcessed)</span></td></tr>
<tr><td style="padding-left:30px !important;">Quarantine cleaned</td><td><span class="value-neutral">$($rep.QuarantineCleaned)</span></td></tr>
<tr><td style="padding-left:30px !important;">Empty folders cleaned</td><td><span class="value-neutral">$($rep.EmptyFoldersCleaned)</span></td></tr>
"@
    }
}
)
</table>

<h3>SOURCES CONFIGURATION</h3>
<table>
<tr><th>Source</th><th>Path / Pattern</th></tr>
$($userProfileSources | Where-Object { $_.Enabled } | ForEach-Object {
    $qStatus = if ($_.EnableQuarantine -eq $false) { 'quarantine disabled' } else { 'quarantine enabled' }
    "<tr><td>$($_.Name)</td><td>$($_.ProfileRoot) (pattern: $($_.PatternProfile), $qStatus)</td></tr>"
})
</table>

<p style="margin:5px 0 0 0;"><strong>Log file location:</strong> $LogFile</p>
</body>
</html>
"@

# Save summary HTML report to logs
$htmlBody | Out-File -FilePath $summaryHtmlFile -Encoding UTF8
Write-Log "Summary HTML report saved: $summaryHtmlFile" -Level "INFO"

# HTML report with detailed table (if there are entries)
if ($executionDetails.Count -gt 0) {
    $htmlRows = ""
    foreach ($item in $executionDetails) {
        $statusColor = switch ($item.Status) {
            "SUCCESS" { "#4CAF50" }  # green
            "FAILED" { "#f44336" }  # red
            "TEST" { "#ff9800" }  # orange
            "REPORTED" { "#2196F3" }  # blue
            default { "#ffffff" }
        }
        if ($item.ProfileType -eq "EmptyFolder") {
            $statusColor = "#9E9E9E"  # grey for empty folders
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

Write-Log "Log saved: $LogFile" -Level "INFO"

# ========== SEND EMAIL ==========
if ($MailSettings.Enabled) {
    Send-ReportEmail -HtmlReportFile $htmlReportFile -HtmlBody $htmlBody -Stats $stats -TestMode $TestMode
}

if ($TestMode) {
    Write-Log "TESTING COMPLETED. Check the report before real execution." -Level "INFO"
}