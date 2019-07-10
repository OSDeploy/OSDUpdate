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
    Version:        19.6.25.0
#>
#======================================================================================
#   Validate Admin Rights
#======================================================================================
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
#   Start Transcript
#======================================================================================
Start-Transcript
#======================================================================================
#   Start Script
#======================================================================================
Write-Host "$PSCommandPath" -ForegroundColor Green
$OSDUpdatePath = (get-item $PSScriptRoot ).FullName
Write-Host "OSDUpdate Path: $OSDUpdatePath" -ForegroundColor Cyan
#======================================================================================
#   Get Child Scripts
#======================================================================================
$OSDScripts = Get-ChildItem $OSDUpdatePath Install-OSDUpdatePackage.ps1 -Recurse | Select-Object -Property *
#======================================================================================
#   Process Child Scripts
#======================================================================================
foreach ($OSDScript in $OSDScripts) {
    Write-Host '========================================================================================' -ForegroundColor DarkGray
    Write-Host "Installing '$($OSDScript.FullName)'" -ForegroundColor Green
    Invoke-Expression "& '$($OSDScript.FullName)'"
}
#======================================================================================
#   Complete
#======================================================================================
Write-Host '========================================================================================' -ForegroundColor DarkGray
Write-Host (Join-Path $PSScriptRoot $MyInvocation.MyCommand.Name) " Complete" -ForegroundColor Green
Write-Host '========================================================================================' -ForegroundColor DarkGray
Stop-Transcript
Start-Sleep 5
#[void](Read-Host 'Press Enter to Continue')