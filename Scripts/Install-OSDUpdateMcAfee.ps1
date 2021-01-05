#Requires -Version 5

<#
.SYNOPSIS
    Installs Updates in an OSDUpdate Package from Child Directories

.DESCRIPTION
    Installs Updates in an OSDUpdate Package from Child Directories

.NOTES
    Author:         David Segura
    Website:        osdeploy.com
    Twitter:        @SeguraOSD
    Version:        20.1.5.1
    This script is only configured to install SuperDAT v2 renamed to xdat.exe
#>
#======================================================================================
#   Begin
#======================================================================================
Write-Host "OSDUpdate McAfee xDAT" -ForegroundColor Green
#======================================================================================
#   Validate
#======================================================================================
$Software = "McAfee VirusScan"
$Installed = Get-ItemProperty ('HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\*','HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*') -EA SilentlyContinue | Where-Object { $_.DisplayName -like "*$Software*" }
#======================================================================================
#   Execute
#======================================================================================
If ($null -eq $Installed) {
    Write-Warning "McAfee VirusScan is not installed"
} else {
    Write-Host "Installing McAfee xDAT" -ForegroundColor Cyan
    if (!(Test-Path "$env:Temp\McAfee\xDAT")) {
        New-Item "$env:Temp\McAfee\xDAT" -ItemType Directory -Force | Out-Null
    }
    Copy-Item "$PSScriptRoot\xdat.exe" "$env:Temp\McAfee\xDAT" -Force | Out-Null
    Start-Process "$env:Temp\McAfee\xDAT\xdat.exe" -ArgumentList '/SILENT','/F' -Wait
}
#======================================================================================
#   Complete
#======================================================================================