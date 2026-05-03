@echo off
:: ----------------------------------------------------------------------
:: Apply-StartLayout.bat
:: Wrapper script that invokes Apply-StartLayout.ps1 in memory to bypass
:: PowerShell execution policy restrictions. Designed for GPO logon.
:: ----------------------------------------------------------------------

:: Set the full path to the PowerShell script (located in the same folder)
set "SCRIPT_PATH=%~dp0Apply-StartLayout.ps1"

:: Set the XML layout file path and pass it to PowerShell via environment variable
set "XmlPath=%~dp0LayoutModification.xml"

:: Run PowerShell, load the script content, and execute it in memory
powershell.exe -Command "$env:XmlPath='%XmlPath%'; $script = Get-Content -Raw '%SCRIPT_PATH%'; Invoke-Expression $script"