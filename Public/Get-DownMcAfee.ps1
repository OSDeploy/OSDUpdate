<#
.SYNOPSIS
Downloads McAfee SuperDATs and Tools

.DESCRIPTION
Downloads McAfee SuperDATs and Tools

.LINK
https://www.osdeploy.com/osdupdate/docs/functions/get-downmcafee

.PARAMETER Product
The type of update to download

.PARAMETER DownloadPath
This is the path to download the updates.  If this is not specified, the Desktop is used

.PARAMETER EPO
Downloads the ZIP version for use in McAfee EPO

.PARAMETER RenameDAT
Renames the DAT without the Version information.  Useful for static installation scripts
#>
function Get-DownMcAfee {
    [CmdletBinding()]
    PARAM (
        [Parameter(Mandatory)]
        [ValidateSet('SuperDAT v2','SuperDAT v3','GetSusp32','GetSusp64','Stinger32','Stinger64')]
        [string]$Download,
        [string]$DownloadPath,
        [switch]$EPO,
        [switch]$RenameDAT
    )
    #===================================================================================================
    #   Paths
    #===================================================================================================
    if (!($DownloadPath)) {$DownloadPath = [Environment]::GetFolderPath("Desktop")}
    if (!(Test-Path "$DownloadPath")) {New-Item -Path "$DownloadPath" -ItemType Directory -Force | Out-Null}

    if ($Download -eq 'GetSusp32') {
        if ($EPO.IsPresent) {
            $DownloadUrl = 'http://downloadcenter.mcafee.com/products/mcafee-avert/getsusp/getsusp-epo.zip'
        } else {
            $DownloadUrl = 'http://downloadcenter.mcafee.com/products/mcafee-avert/getsusp/getsusp.exe'
        }
    }

    if ($Download -eq 'GetSusp64') {
        if ($EPO.IsPresent) {
            $DownloadUrl = 'http://downloadcenter.mcafee.com/products/mcafee-avert/getsusp/getsusp64-epo.zip'
        } else {
            $DownloadUrl = 'http://downloadcenter.mcafee.com/products/mcafee-avert/getsusp/getsusp64.exe'
        }
    }

    if ($Download -eq 'Stinger32') {
        if ($EPO.IsPresent) {
            $DownloadUrl = 'http://downloadcenter.mcafee.com/products/mcafee-avert/Stinger/stinger32-epo.zip'
        } else {
            $DownloadUrl = 'http://downloadcenter.mcafee.com/products/mcafee-avert/Stinger/stinger32.exe'
        }
    }

    if ($Download -eq 'Stinger64') {
        if ($EPO.IsPresent) {
            $DownloadUrl = 'http://downloadcenter.mcafee.com/products/mcafee-avert/Stinger/stinger64-epo.zip'
        } else {
            $DownloadUrl = 'http://downloadcenter.mcafee.com/products/mcafee-avert/Stinger/stinger64.exe'
        }
    }

    if ($Download -eq 'SuperDAT v2') {
        if ($EPO.IsPresent) {
            $DownloadString = 'download.nai.com/products/DatFiles/4.x/NAI/avvepo'
            $link = (Invoke-WebRequest -Uri 'https://www.mcafee.com/enterprise/en-us/downloads/security-updates.html').Links | Where-Object {$_.href -like "*$DownloadString*"}
            $DownloadUrl = $link.href
        } else {
            $DownloadString = 'download.nai.com/products/licensed/superdat/english/intel'
            $link = (Invoke-WebRequest -Uri 'https://www.mcafee.com/enterprise/en-us/downloads/security-updates.html').Links | Where-Object {$_.href -like "*$DownloadString*"}
            $DownloadUrl = $link.href
        }
    }

    if ($Download -eq 'SuperDAT v3') {
        if ($EPO.IsPresent) {
            $DownloadString = 'download.nai.com/products/datfiles/DATv3/epoV3'
            $link = (Invoke-WebRequest -Uri 'https://www.mcafee.com/enterprise/en-us/downloads/security-updates.html').Links | Where-Object {$_.href -like "*$DownloadString*"}
            $DownloadUrl = $link.href
        } else {
            $DownloadString = 'download.nai.com/products/datfiles/DATv3/V3'
            $link = (Invoke-WebRequest -Uri 'https://www.mcafee.com/enterprise/en-us/downloads/security-updates.html').Links | Where-Object {$_.href -like "*$DownloadString*"}
            $DownloadUrl = $link.href
        }
    }
    #===================================================================================================
    #   Download
    #===================================================================================================
    Write-Host "DownloadUrl: $DownloadUrl" -ForegroundColor Cyan
    Write-Host "DownloadPath: $DownloadPath" -ForegroundColor Cyan
    Write-Host "Product: $Download" -ForegroundColor Cyan

    if ($Download -eq 'SuperDAT v2' -and $RenameDAT.IsPresent -and (!($EPO.IsPresent)) ) {
        Write-Host "RenameDAT: $DownloadPath\xdat.exe" -ForegroundColor Cyan
        Start-BitsTransfer -Source $DownloadUrl -Destination "$DownloadPath\xdat.exe"
    } elseif ($Download -eq 'SuperDAT v3' -and $RenameDAT.IsPresent -and (!($EPO.IsPresent)) ) {
        Write-Host "RenameDAT: $DownloadPath\DATv3.exe" -ForegroundColor Cyan
        Start-BitsTransfer -Source $DownloadUrl -Destination "$DownloadPath\DATv3.exe"
    } else {
        Start-BitsTransfer -Source $DownloadUrl -Destination "$DownloadPath"
    }
}