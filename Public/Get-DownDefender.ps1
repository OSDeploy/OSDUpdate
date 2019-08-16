<#
.SYNOPSIS
Downloads the latest Microsoft Defender Definition Updates

.DESCRIPTION
Downloads the latest Microsoft Defender Definition Updates

.LINK
https://osdupdate.osdeploy.com/module/functions/get-downdefender

.PARAMETER OS
Self Explanatory

.PARAMETER OSArch
Windows OS Architecture

.PARAMETER DownloadPath
This is the path to download the updates.  If this is not specified, the Desktop is used
#>
function Get-DownDefender {
    [CmdletBinding()]
    PARAM (
        [string]$DownloadPath,

        [Parameter(Mandatory)]
        [ValidateSet('Windows 8-10','Windows V-7')]
        [string]$OS,

        [Parameter(Mandatory)]
        [ValidateSet('32-Bit','64-Bit')]
        [string]$OSArch
    )
    #===================================================================================================
    #   Paths
    #===================================================================================================
    if (!($DownloadPath)) {$DownloadPath = [Environment]::GetFolderPath("Desktop")}
    if (!(Test-Path "$DownloadPath")) {New-Item -Path "$DownloadPath" -ItemType Directory -Force | Out-Null}
    
    Write-Host "DownloadPath: $DownloadPath" -ForegroundColor Cyan
    Write-Host "OS: $OS" -ForegroundColor Cyan
    Write-Host "Arch: $OSArch" -ForegroundColor Cyan

    if ($OS -eq 'Windows V-7' -and $OSArch -eq '32-Bit') {
        $DownloadUrl = 'https://go.microsoft.com/fwlink/?LinkID=121721&clcid=0x409&arch=x86&eng=0.0.0.0&avdelta=0.0.0.0&asdelta=0.0.0.0&prod=925A3ACA-C353-458A-AC8D-A7E5EB378092'
    }

    if ($OS -eq 'Windows V-7' -and $OSArch -eq '64-Bit') {
        $DownloadUrl = 'https://go.microsoft.com/fwlink/?LinkID=121721&clcid=0x409&arch=x86&eng=0.0.0.0&avdelta=0.0.0.0&asdelta=0.0.0.0&prod=925A3ACA-C353-458A-AC8D-A7E5EB378092'
    }

    if ($OS -eq 'Windows 8-10' -and $OSArch -eq '32-Bit') {
        $DownloadUrl = 'https://go.microsoft.com/fwlink/?LinkID=121721&arch=x86'
    }

    if ($OS -eq 'Windows 8-10' -and $OSArch -eq '64-Bit') {
        $DownloadUrl = 'https://go.microsoft.com/fwlink/?LinkID=121721&arch=x64'
    }
    #===================================================================================================
    #   Download
    #===================================================================================================
    Write-Host "DownloadUrl: $DownloadUrl" -ForegroundColor Cyan
    Write-Host "DownloadPath: $DownloadPath" -ForegroundColor Cyan
    Invoke-WebRequest -Uri $DownloadUrl -OutFile "$DownloadPath\mpam-fe $OS $OSArch.exe"
}