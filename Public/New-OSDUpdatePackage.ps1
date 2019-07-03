<#
.SYNOPSIS
Creates Package Bundles of Microsoft Updates

.DESCRIPTION
Creates Package Bundles of Microsoft Updates
Requires BITS for downloading the updates
Requires Internet access for downloading the updates

.LINK
https://www.osdeploy.com/osdupdate/functions/new-osdupdatepackage

.PARAMETER PackagePath
Package Content will be downloaded into the PackagePath

.PARAMETER PackageName
The name of the OSDUpdate Package.  These values are predefined

.PARAMETER AppendPackageName
Downloads Updates in the PackagePath in a PackageName subdirectory

.PARAMETER GridView
Displays the results in GridView with -PassThru.  Updates selected in GridView can be selected

.PARAMETER HideDownloaded
Hides downloaded updates from the results

.PARAMETER OfficeProfile
Downloads Office Updates with the selected Profile

.PARAMETER OfficeSetupUpdatesPath
Updates Directory for Office Setup

.PARAMETER RemoveSuperseded
Remove Superseded Updates

.PARAMETER SkipInstallScript
Skips adding $PackagePath\Install-OSDUpdatePackage.ps1

.PARAMETER SkipUpdateScript
Skips adding $PackagePath\Update-OSDUpdatePackage.ps1

#>
function New-OSDUpdatePackage {
    [CmdletBinding()]
    PARAM (
        [Parameter(Mandatory = $True)]
        [string]$PackagePath,

        [Parameter(Mandatory = $True)]
        [ValidateSet(
            #================================
            #   Office
            #================================
            'Office 2010 32-Bit',
            'Office 2010 64-Bit',
            'Office 2013 32-Bit',
            'Office 2013 64-Bit',
            'Office 2016 32-Bit',
            'Office 2016 64-Bit',
            #================================
            #   Windows
            #================================
            'Windows 7 x64',
            'Windows 7 x86',
            'Windows 10 x64 1803',
            'Windows 10 x64 1809',
            'Windows 10 x64 1903',
            'Windows 10 x86 1803',
            'Windows 10 x86 1809',
            'Windows 10 x86 1903',
            'Windows Server 2016 1607',
            'Windows Server 2016 1709',
            'Windows Server 2016 1803',
            'Windows Server 2019 1809',
            #================================
            #   Other
            #================================
            'Servicing Stacks')]
        [string]$PackageName,

        [switch]$AppendPackageName,
        [switch]$GridView,
        [switch]$HideDownloaded,

        [ValidateSet('Default','Proofing','Language','All')]
        [string]$OfficeProfile = 'Default',

        [string]$OfficeSetupUpdatesPath,

        [switch]$RemoveSuperseded,
        [switch]$SkipInstallScript,
        [switch]$SkipUpdateScript

    )
    #===================================================================================================
    #   AppendPackageName
    #===================================================================================================
    if ($AppendPackageName) {
        $PackagePath = "$PackagePath\$PackageName"
    }
    #===================================================================================================
    #   PackagePath
    #===================================================================================================
    if (!(Test-Path "$PackagePath")) {New-Item -Path "$PackagePath" -ItemType Directory -Force | Out-Null}
    #===================================================================================================
    #   Get-OSDUpdate
    #===================================================================================================
    $OSDUpdate = @()
    $OSDUpdate = Get-OSDUpdate
    Write-Host $PackageName -ForegroundColor Green
    #===================================================================================================
    #   Filter Catalog
    #===================================================================================================
    if ($PackageName -match 'Office') {
        $OSDUpdate = $OSDUpdate | Where-Object {$_.Catalog -eq $PackageName}
    }

    if ($PackageName -like "Windows*") {
        if ($PackageName -like "Windows 7*") {
            $OSDUpdate = $OSDUpdate | Where-Object {$_.Catalog -eq 'Windows 7'}
        }
        if ($PackageName -like "Windows 10*") {
            $OSDUpdate = $OSDUpdate | Where-Object {$_.Catalog -eq 'Windows 10'}
        }
        if ($PackageName -like "Windows Server 2016*") {
            $OSDUpdate = $OSDUpdate | Where-Object {$_.Catalog -eq 'Windows Server 2016'}
        }
        if ($PackageName -like "Windows Server 2019*") {
            $OSDUpdate = $OSDUpdate | Where-Object {$_.Catalog -eq 'Windows Server 2019'}
        }
    }
    #===================================================================================================
    #   AllOSDUpdates
    #===================================================================================================
    $AllOSDUpdates = $OSDUpdate
    #===================================================================================================
    #   Office Superseded Updates
    #===================================================================================================
    if ($PackageName -match 'Office') {
        $OSDUpdate = $OSDUpdate | Sort-Object OriginUri -Unique
        $OSDUpdate = $OSDUpdate | Sort-Object CreationDate -Descending

        $CurrentUpdates = @()
        $SupersededUpdates = @()

        foreach ($OfficeUpdate in $OSDUpdate) {
            $SkipUpdate = $false

            foreach ($CurrentUpdate in $CurrentUpdates) {
                if ($($OfficeUpdate.FileName) -eq $($CurrentUpdate.FileName)) {$SkipUpdate = $true}
            }

            if ($SkipUpdate) {
                $SupersededUpdates += $OfficeUpdate
            } else {
                $CurrentUpdates += $OfficeUpdate
            }
        }
        $OSDUpdate = $CurrentUpdates
    }
    #===================================================================================================
    #   Find Existing Updates
    #===================================================================================================
    $LocalUpdates = @()
    $LocalSuperseded = @()

    $LocalUpdates = Get-ChildItem -Path "$PackagePath\*" -Directory -Recurse | Select-Object -Property *
    $LocalUpdatesMsp = Get-ChildItem -Path "$PackagePath\*" *.msp -File -Recurse | Select-Object -Property * | Sort-Object CreationDate -Descending
    $LocalUpdatesXml = Get-ChildItem -Path "$PackagePath\*" *.xml -File -Recurse | Select-Object -Property * | Sort-Object CreationDate -Descending

    foreach ($Update in $LocalUpdates) {
        if ($CurrentUpdates.Title -NotContains $Update.Name) {$LocalSuperseded += $Update.FullName}
    }
    #===================================================================================================
    #   Remove Superseded Update Directories
    #===================================================================================================
    foreach ($Update in $LocalSuperseded) {
        if ($RemoveSuperseded.IsPresent) {
            Write-Warning "Removing Superseded: $Update"
            Remove-Item $Update -Recurse -Force | Out-Null
        } else {
            Write-Warning "Superseded: $Update"
        }
    }
    #===================================================================================================
    #   Remove Superseded Update Files
    #===================================================================================================
    if ($PackageName -match 'Office') {
        foreach ($Update in $SupersededUpdates) {
            $SupersededMsp = "$PackagePath\$($Update.Title)\$([IO.Path]::GetFileNameWithoutExtension($Update.FileName)).msp"
            $SupersededXml = "$PackagePath\$($Update.Title)\$([IO.Path]::GetFileNameWithoutExtension($Update.FileName)).xml"
            if (Test-Path "$SupersededMsp") {
                if ($RemoveSuperseded.IsPresent) {
                    Write-Warning "Removing Superseded: $SupersededMsp"
                    Remove-Item $SupersededMsp -Force | Out-Null
                } else {
                    Write-Warning "Superseded: $SupersededMsp"
                }
            }
            if (Test-Path "$SupersededXml") {
                if ($RemoveSuperseded.IsPresent) {
                    Write-Warning "Removing Superseded: $SupersededXml"
                    Remove-Item $SupersededXml -Force | Out-Null
                } else {
                    Write-Warning "Superseded: $SupersededXml"
                }
            }
        }
    }
    #===================================================================================================
    #   Get Downloaded Updates
    #===================================================================================================
    foreach ($Update in $OSDUpdate) {
        if ($PackageName -like "Windows*" -or $PackageName -eq 'Servicing Stacks') {
            $FullUpdatePath = "$PackagePath\$($Update.Title)\$($Update.FileName)"
            if (Test-Path $FullUpdatePath) {
                $Update.OSDStatus = 'Downloaded'
            }
        }
        if ($PackageName -match 'Office') {
            $FullUpdatePath = "$PackagePath\$($Update.Title)\$([IO.Path]::GetFileNameWithoutExtension($Update.FileName)).msp"
            if (Test-Path $FullUpdatePath) {
                $Update.OSDStatus = 'Downloaded'
            }
        }
    }
    #===================================================================================================
    #   Filter Other
    #===================================================================================================
    if ($PackageName -eq 'Servicing Stacks') {
        $OSDUpdate = $OSDUpdate | Where-Object {$_.UpdateGroup -eq 'SSU'}
    }

    if ($PackageName -like "Windows*") {
        if ($PackageName -like "*x64*") {$OSDUpdate = $OSDUpdate | Where-Object {$_.UpdateArch -eq 'x64'}}
        if ($PackageName -like "*x86*") {$OSDUpdate = $OSDUpdate | Where-Object {$_.UpdateArch -eq 'x86'}}
        if ($PackageName -like "*1507*") {$OSDUpdate = $OSDUpdate | Where-Object {$_.UpdateBuild -eq '1507'}}
        if ($PackageName -like "*1511*") {$OSDUpdate = $OSDUpdate | Where-Object {$_.UpdateBuild -eq '1511'}}
        if ($PackageName -like "*1607*") {$OSDUpdate = $OSDUpdate | Where-Object {$_.UpdateBuild -eq '1607'}}
        if ($PackageName -like "*1703*") {$OSDUpdate = $OSDUpdate | Where-Object {$_.UpdateBuild -eq '1703'}}
        if ($PackageName -like "*1709*") {$OSDUpdate = $OSDUpdate | Where-Object {$_.UpdateBuild -eq '1709'}}
        if ($PackageName -like "*1803*") {$OSDUpdate = $OSDUpdate | Where-Object {$_.UpdateBuild -eq '1803'}}
        if ($PackageName -like "*1809*") {$OSDUpdate = $OSDUpdate | Where-Object {$_.UpdateBuild -eq '1809'}}
        if ($PackageName -like "*1903*") {$OSDUpdate = $OSDUpdate | Where-Object {$_.UpdateBuild -eq '1903'}}
    }
    $OSDUpdate | Export-Clixml -Path "$PackagePath\OSDUpdatePackage.xml" -Force | Out-Null

    if ($PackageName -like "Office*") {
        if ($OfficeProfile -eq 'Default') {
            $OSDUpdate = $OSDUpdate | Where-Object {$_.FileName -like "*none*" -or $_.FileName -like "*en-us*"}
            $OSDUpdate = $OSDUpdate | Where-Object {$_.Title -notlike "*Language Pack*"}
        }
        if ($OfficeProfile -eq 'Language') {
            $OSDUpdate = $OSDUpdate | Where-Object {$_.FileName -notlike "*none*" -and $_.FileName -notlike "*en-us*"}
        }
        if ($OfficeProfile -eq 'Proofing') {
            $OSDUpdate = $OSDUpdate | Where-Object {$_.FileName -like "*Proof*"}
        }
    }
    #===================================================================================================
    #   HideDownloaded
    #===================================================================================================
    if ($HideDownloaded.IsPresent) {
        $OSDUpdate = $OSDUpdate | Where-Object {$_.OSDStatus -ne 'Downloaded'}
    }
    #===================================================================================================
    #   GridView
    #===================================================================================================
    if ($PackageName -like "Office*") {
        $OSDUpdate = $OSDUpdate | Select-Object -Property OSDStatus,Catalog,CreationDate,KBNumber,Title,FileName,Size,FileUri,OriginUri,OSDGuid
    } else {
        $OSDUpdate = $OSDUpdate | Select-Object -Property OSDStatus,Catalog,UpdateOS,UpdateArch,UpdateBuild,CreationDate,KBNumber,Title,FileName,Size,FileUri,OriginUri,OSDGuid
    }
    if ($GridView.IsPresent) {$OSDUpdate = $OSDUpdate | Out-GridView -PassThru -Title "Select OSDUpdate Downloads to include in the Package"}
    #===================================================================================================
    #   Sort
    #===================================================================================================
    $OSDUpdate = $OSDUpdate | Sort-Object DateCreated
    #===================================================================================================
    #   Office Download
    #===================================================================================================
    if ($PackageName -like "Office*") {
        foreach ($Update in $OSDUpdate) {
            $UpdateFile = $($Update.FileName)
            $MspFile = $UpdateFile -replace '.cab', '.msp'
            $DownloadDirectory = "$PackagePath\$($Update.Title)"

            if (!(Test-Path "$DownloadDirectory")) {New-Item -Path "$DownloadDirectory" -ItemType Directory -Force | Out-Null}
        
            if (Test-Path "$DownloadDirectory\$MspFile") {
                Write-Host "$DownloadDirectory\$MspFile" -ForegroundColor Cyan
            } else {
                Write-Host "$DownloadDirectory\$MspFile" -ForegroundColor Cyan
                Write-Host "Download: $($Update.OriginUri)" -ForegroundColor Gray
                Start-BitsTransfer -Source $($Update.OriginUri) -Destination "$DownloadDirectory\$UpdateFile"
            }

            if ((Test-Path "$DownloadDirectory\$UpdateFile") -and (!(Test-Path "$DownloadDirectory\$MspFile"))) {
                Write-Host "Expand: $DownloadDirectory\$MspFile" -ForegroundColor Gray
                expand "$DownloadDirectory\$UpdateFile" -F:* "$DownloadDirectory" | Out-Null
            }

            if ((Test-Path "$DownloadDirectory\$UpdateFile") -and (Test-Path "$DownloadDirectory\$MspFile")) {
                Write-Host "Remove: $DownloadDirectory\$UpdateFile" -ForegroundColor Gray
                Remove-Item "$DownloadDirectory\$UpdateFile" -Force | Out-Null
            }
            #===================================================================================================
            #   Office Setup Updates
            #===================================================================================================
            if ($OfficeSetupUpdatesPath) {
                if (!(Test-Path "$OfficeSetupUpdatesPath")) {New-Item -Path "$OfficeSetupUpdatesPath" -ItemType Directory -Force | Out-Null}
                Write-Host "Date Created: $($Update.CreationDate)" -ForegroundColor Gray
                Write-Host "Source: $DownloadDirectory\$MspFile" -ForegroundColor Gray
                Write-Host "Destination: $OfficeSetupUpdatesPath\$MspFile" -ForegroundColor Gray
                Copy-Item -Path "$DownloadDirectory\$MspFile" "$OfficeSetupUpdatesPath\$MspFile" -Force
            }
        }
    } else {
        foreach ($Update in $OSDUpdate) {
            $UpdateFile = $($Update.FileName)
            $DownloadDirectory = "$PackagePath\$($Update.Title)"

            if (!(Test-Path "$DownloadDirectory")) {New-Item -Path "$DownloadDirectory" -ItemType Directory -Force | Out-Null}
        
            if (Test-Path "$DownloadDirectory\$UpdateFile") {
                Write-Host "$($Update.Title)" -ForegroundColor Cyan
                Write-Host "$DownloadDirectory\$UpdateFile" -ForegroundColor Gray
                #Write-Host "Update already downloaded" -ForegroundColor Gray
            } else {
                Write-Host "$($Update.Title)" -ForegroundColor Cyan
                Write-Host "$($Update.OriginUri)" -ForegroundColor Gray
                Write-Host "$DownloadDirectory\$UpdateFile" -ForegroundColor Gray
                Start-BitsTransfer -Source $($Update.OriginUri) -Destination "$DownloadDirectory\$UpdateFile"
            }
        }
    }
    #===================================================================================================
    #   Install Script
    #===================================================================================================
    if (!($SkipInstallScript)) {
        Write-Host "Install Script: $PackagePath\Install-OSDUpdatePackage.ps1" -ForegroundColor Green
        Copy-Item "$($MyInvocation.MyCommand.Module.ModuleBase)\Scripts\Install-OSDUpdatePackage.ps1" "$PackagePath" -Force | Out-Null
    }
    #===================================================================================================
    #   Update Script
    #===================================================================================================
    if (!($SkipUpdateScript)) {
        Write-Host "Update Script: $PackagePath\Update-OSDUpdatePackage.ps1" -ForegroundColor Green
        $ExportLine = "New-OSDUpdatePackage -PackageName '$PackageName' -PackagePath ""`$PSScriptRoot"" -RemoveSuperseded"
        $ExportLine | Out-File -FilePath "$PackagePath\Update-OSDUpdatePackage.ps1"
    }
}