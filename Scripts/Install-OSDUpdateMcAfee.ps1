#Requires -Version 5

<#
.SYNOPSIS
    Installs McAfee SuperDAT v2 renamed to xdat.exe

.DESCRIPTION
    Installs McAfee SuperDAT v2 renamed to xdat.exe

.NOTES
    Author:         David Segura
    Website:        osdeploy.com
    Twitter:        @SeguraOSD
    Version:        21.1.7.2
#>
#======================================================================================
#   Validate Admin Rights
#======================================================================================
Write-Host ""
# Verify Running as Admin
$isAdmin = ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")
If (!( $isAdmin )) {
    Write-Host "Checking User Account Control settings ..." -ForegroundColor Green
    if ((Get-ItemProperty HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System).EnableLUA -eq 0) {
        #UAC Disabled
        Write-Host '========================================================================================' -ForegroundColor DarkGray
        Write-Host "User Account Control is Disabled ... " -ForegroundColor Green
        Write-Host "You will need to correct your UAC Settings ..." -ForegroundColor Green
        Write-Host "Try running this script in an Elevated PowerShell session ... Exiting" -ForegroundColor Green
        Write-Host '========================================================================================' -ForegroundColor DarkGray
        Start-Sleep -s 10
        Exit 0
    } else {
        #UAC Enabled
        Write-Host "UAC is Enabled" -ForegroundColor Green
        Start-Sleep -s 3
        if ($Silent) {
            Write-Host "-- Restarting as Administrator (Silent)" -ForegroundColor Cyan ; Start-Sleep -Seconds 1
            Start-Process powershell.exe "-NoProfile -ExecutionPolicy Bypass -File `"$PSCommandPath`" -Silent" -Verb RunAs -Wait
        } elseif($Restart) {
            Write-Host "-- Restarting as Administrator (Restart)" -ForegroundColor Cyan ; Start-Sleep -Seconds 1
            Start-Process powershell.exe "-NoProfile -ExecutionPolicy Bypass -File `"$PSCommandPath`" -Restart" -Verb RunAs -Wait
        } else {
            Write-Host "-- Restarting as Administrator" -ForegroundColor Cyan ; Start-Sleep -Seconds 1
            Start-Process powershell.exe "-NoProfile -ExecutionPolicy Bypass -File `"$PSCommandPath`"" -Verb RunAs -Wait
        }
        Exit 0
    }
} else {
    Write-Host '========================================================================================' -ForegroundColor DarkGray
    Write-Host "-- Running with Elevated Permissions ..." -ForegroundColor Cyan ; Start-Sleep -Seconds 1
    Write-Host '========================================================================================' -ForegroundColor DarkGray
}
#======================================================================================
#   Script Information
#======================================================================================
$Invocation = (Get-Variable MyInvocation -Scope Script).Value
$ScriptPath = Split-Path -Parent $Invocation.MyCommand.Path
$ParentName = Split-Path $ScriptPath -Leaf
#======================================================================================
#   Logs
#======================================================================================
$OSDAppName = "OSDUpdate-$ParentName"
$OSDLogs = "$env:Temp"
if (!(Test-Path $OSDLogs)) {New-Item $OSDLogs -ItemType Directory -Force | Out-Null}
$OSDLogName = "$((Get-Date).ToString('yyyy-MM-dd-HHmmss'))-$OSDAppName.log"
Start-Transcript -Path (Join-Path $OSDLogs $OSDLogName)
#======================================================================================
#   Start Script
#======================================================================================
Write-Host "Start ... $(Join-Path $PSScriptRoot $MyInvocation.MyCommand.Name)" -ForegroundColor Green
Write-Host ""
#======================================================================================
#   OS Information
#======================================================================================
$OSCaption = $((Get-WmiObject -Class Win32_OperatingSystem).Caption).Trim()
$OSArchitecture = $((Get-WmiObject -Class Win32_OperatingSystem).OSArchitecture).Trim()
$OSProductType = $((Get-WmiObject -Class Win32_OperatingSystem).ProductType)
$OSVersion = $((Get-WmiObject -Class Win32_OperatingSystem).Version).Trim()
$OSBuildNumber = $((Get-WmiObject -Class Win32_OperatingSystem).BuildNumber).Trim()
Write-Host "Operating System: $OSCaption" -ForegroundColor Cyan
Write-Host "OS Architecture: $OSArchitecture" -ForegroundColor Cyan
Write-Host "OS ProductType: $OSProductType" -ForegroundColor Cyan
Write-Host "OS Version: $OSVersion" -ForegroundColor Cyan
Write-Host "OS Build Number: $OSBuildNumber" -ForegroundColor Cyan
if ($OSVersion -Like "10*") {
    $OSReleaseID = (Get-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion" -Name ReleaseId).ReleaseId
    Write-Host "OS Release ID: $OSReleaseID" -ForegroundColor Cyan
}
#======================================================================================
#   Begin
#======================================================================================
Write-Host "Updating McAfee xDAT Signatures" -ForegroundColor Green
#======================================================================================
#   Validate McAfee SuperDAT v2
#======================================================================================
if (Test-Path "$PSScriptRoot\xdat.exe") {
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
}
#======================================================================================
#   Validate McAfee SuperDAT v3
#======================================================================================
if (Test-Path "$PSScriptRoot\V3_xdat.exe") {
	$Software = "McAfee Endpoint Security Threat Prevention"
	$Installed = Get-ItemProperty ('HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\*','HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*') -EA SilentlyContinue | Where-Object { $_.DisplayName -like "*$Software*" }
	#======================================================================================
	#   Execute
	#======================================================================================
	If ($null -eq $Installed) {
		Write-Warning "McAfee Endpoint Security is not installed"
	} else {
		Write-Host "Installing McAfee xDAT" -ForegroundColor Cyan
		if (!(Test-Path "$env:Temp\McAfee\xDAT")) {
			New-Item "$env:Temp\McAfee\xDAT" -ItemType Directory -Force | Out-Null
		}
		Copy-Item "$PSScriptRoot\V3_xdat.exe" "$env:Temp\McAfee\xDAT" -Force | Out-Null
		Start-Process "$env:Temp\McAfee\xDAT\V3_xdat.exe" -ArgumentList '/SILENT','/F' -Wait
	}
}
#======================================================================================
#   Complete
#======================================================================================
Write-Host ""
Write-Host "Complete ... $(Join-Path $PSScriptRoot $MyInvocation.MyCommand.Name)" -ForegroundColor Green
Stop-Transcript
Start-Sleep 5