<#
.SYNOPSIS
Downloads Current Microsoft Updates

.DESCRIPTION
Downloads Current Microsoft Updates
Requires BITS for downloading the updates
Requires Internet access for downloading the updates

.LINK
https://www.osdeploy.com/osdupdate/docs/functions/get-downosdupdate

.PARAMETER CatalogOffice
The Microsoft Office Update Catalog selected for Updates

.PARAMETER OfficeProfile
Microsoft Office Update Type

.PARAMETER OfficeSetupUpdatesPath
This is the Updates directory in your Office installation source
MSP files will be extracted from the DownloadsPath
This directory should be cleared first

.PARAMETER RemoveSuperseded
Removes the Superseded Microsoft Office Updates from the Repository

.PARAMETER CatalogWindows
The Microsoft Windows Update Catalog selected for Updates

.PARAMETER UpdateArch
Architecture of the Update

.PARAMETER UpdateBuild
Windows Build for the Update

.PARAMETER UpdateGroup
Windows Update Type

.PARAMETER GridView
Displays the results in GridView with -PassThru

.PARAMETER RepositoryRootPath
Full Path of the OSDUpdate Repository
#>

function Get-DownOSDUpdate {
    [CmdletBinding(DefaultParameterSetName = 'Office')]
    PARAM (
        #===================================================================================================
        #   Windows and Office
        #===================================================================================================
        #[Parameter(ParameterSetName = 'Windows')]
        #[Parameter(ParameterSetName = 'Office')]
        [switch]$AddInstallScript,
        [Parameter(ParameterSetName = 'Windows')]
        [Parameter(ParameterSetName = 'Office')]
        [switch]$GridView,

        [switch]$RemoveSuperseded,
        [Parameter(Mandatory = $True)]
        [string]$RepositoryRootPath,
        #===================================================================================================
        #   Windows Tab Only
        #===================================================================================================
        [Parameter(ParameterSetName = 'Windows', Mandatory = $True)]
        [ValidateSet(
            'Windows 7',
            #'Windows 8.1',
            #'Windows 8.1 Dynamic Update',
            'Windows 10',
            'Windows 10 Dynamic Update',
            'Windows 10 Feature On Demand',
            'Windows 10 Language Packs',
            'Windows 10 Language Interface Packs',
            'Windows Server 2012 R2',
            'Windows Server 2012 R2 Dynamic Update',
            'Windows Server 2016',
            'Windows Server 2019')]
        [string]$CatalogWindows,

        [Parameter (ParameterSetName = 'Windows')]
        [ValidateSet ('x64','x86')]
        [string]$UpdateArch,

        [Parameter (ParameterSetName = 'Windows')]
        [ValidateSet (1903,1809,1803,1709,1703,1607,1511,1507)]
        [string]$UpdateBuild,

        [Parameter(ParameterSetName = 'Windows')]
        [ValidateSet(
            'Setup Dynamic Update',
            'Component Dynamic Update',
            'Adobe Flash Player',
            'DotNet Framework',
            'Latest Cumulative Update LCU',
            'Servicing Stack Update SSU')]
        [string]$UpdateGroup,
        #===================================================================================================
        #   Office Tab Only
        #===================================================================================================
        [Parameter(ParameterSetName = 'Office', Mandatory = $True)]
        [ValidateSet(
            'Office 2010 32-Bit',
            'Office 2010 64-Bit',
            'Office 2013 32-Bit',
            'Office 2013 64-Bit',
            'Office 2016 32-Bit',
            'Office 2016 64-Bit')]
        [string]$CatalogOffice,

        [Parameter(ParameterSetName = 'Office', Mandatory = $True)]
        [ValidateSet(
            'Default',
            'Proofing',
            'Language',
            'All')]
        [string]$OfficeProfile,

        [Parameter(ParameterSetName = 'Office')]
        [string]$OfficeSetupUpdatesPath
    )

    #===================================================================================================
    #   Variables
    #===================================================================================================
    $AllOSDUpdates = @()
    #===================================================================================================
    #   Repository
    #===================================================================================================
    if ($CatalogOffice) {
        New-OSDUpdateRepository -RepositoryRootPath "$RepositoryRootPath" -Catalog $CatalogOffice
        $DownloadsPath = "$RepositoryRootPath\$CatalogOffice"
    }
    if ($CatalogWindows) {
        New-OSDUpdateRepository -RepositoryRootPath "$RepositoryRootPath" -Catalog $CatalogWindows
        $DownloadsPath = "$RepositoryRootPath\$CatalogWindows"
        Write-Verbose $DownloadsPath
    }
    #===================================================================================================
    #   CatalogOffice
    #===================================================================================================
    if ($CatalogOffice) {
        $AllDownOSDUpdate = Get-OSDUpdate -CatalogOffice $CatalogOffice -OfficeProfile All -Silent
        if ($OfficeProfile) {
            $AllOSDUpdates = Get-OSDUpdate -CatalogOffice $CatalogOffice -OfficeProfile $OfficeProfile
        } else {
            $AllOSDUpdates = Get-OSDUpdate -CatalogOffice $CatalogOffice
        }
    }
    #===================================================================================================
    #   Catalog Windows
    #===================================================================================================
    if ($CatalogWindows) {
        $AllDownOSDUpdate = Get-OSDUpdate -CatalogWindows $CatalogWindows
        $AllOSDUpdates = $AllDownOSDUpdate
    }
    #===================================================================================================
    #   Existing Updates
    #===================================================================================================
    $ExistingUpdates = @()
    $SupersededUpdates = @()

    $ExistingUpdates = Get-ChildItem -Path "$DownloadsPath\*" -Directory -Recurse | Select-Object -Property *

    foreach ($Update in $ExistingUpdates) {
        if ($AllDownOSDUpdate.Title -NotContains $Update.Name) {$SupersededUpdates += $Update.FullName}
    }
    #===================================================================================================
    #   Superseded Updates
    #===================================================================================================
    foreach ($Update in $SupersededUpdates) {
        if ($RemoveSuperseded.IsPresent) {
            Write-Warning "Removing Superseded: $Update"
            Remove-Item $Update -Recurse -Force | Out-Null
        } else {
            Write-Warning "Superseded: $Update"
        }
    }
    #===================================================================================================
    #   Get Downloaded Updates
    #===================================================================================================
    foreach ($Update in $AllOSDUpdates) {
        if ($CatalogWindows) {
            $FullUpdatePath = "$RepositoryRootPath\$CatalogWindows\$($Update.Title)\$($Update.FileName)"
            if (Test-Path $FullUpdatePath) {
                $Update.OSDStatus = "Downloaded"
            }
        }
    }
    #===================================================================================================
    #   UpdateArch
    #===================================================================================================
    if ($UpdateArch -eq 'x64') {$AllOSDUpdates = $AllOSDUpdates | Where-Object {$_.UpdateArch -eq 'x64'}}
    if ($UpdateArch -eq 'x86') {$AllOSDUpdates = $AllOSDUpdates | Where-Object {$_.UpdateArch -eq 'x86'}}
    #===================================================================================================
    #   Update Build
    #===================================================================================================
    if ($UpdateBuild -eq 1903) {$AllOSDUpdates = $AllOSDUpdates | Where-Object {$_.UpdateBuild -eq '1903'}}
    if ($UpdateBuild -eq 1809) {$AllOSDUpdates = $AllOSDUpdates | Where-Object {$_.UpdateBuild -eq '1809'}}
    if ($UpdateBuild -eq 1803) {$AllOSDUpdates = $AllOSDUpdates | Where-Object {$_.UpdateBuild -eq '1803'}}
    if ($UpdateBuild -eq 1709) {$AllOSDUpdates = $AllOSDUpdates | Where-Object {$_.UpdateBuild -eq '1709'}}
    if ($UpdateBuild -eq 1703) {$AllOSDUpdates = $AllOSDUpdates | Where-Object {$_.UpdateBuild -eq '1703'}}
    if ($UpdateBuild -eq 1607) {$AllOSDUpdates = $AllOSDUpdates | Where-Object {$_.UpdateBuild -eq '1607'}}
    if ($UpdateBuild -eq 1511) {$AllOSDUpdates = $AllOSDUpdates | Where-Object {$_.UpdateBuild -eq '1511'}}
    if ($UpdateBuild -eq 1507) {$AllOSDUpdates = $AllOSDUpdates | Where-Object {$_.UpdateBuild -eq '1507'}}
    #===================================================================================================
    #   UpdateGroup
    #===================================================================================================
    if ($UpdateGroup -eq 'Setup Dynamic Update') {$AllOSDUpdates = $AllOSDUpdates | Where-Object {$_.UpdateGroup -eq 'SetupDU'}}
    if ($UpdateGroup -eq 'Component Dynamic Update') {$AllOSDUpdates = $AllOSDUpdates | Where-Object {$_.UpdateGroup -like "ComponentDU*"}}
    if ($UpdateGroup -eq 'Adobe Flash Player') {$AllOSDUpdates = $AllOSDUpdates | Where-Object {$_.UpdateGroup -eq 'AdobeSU'}}
    if ($UpdateGroup -eq 'DotNet Framework') {$AllOSDUpdates = $AllOSDUpdates | Where-Object {$_.UpdateGroup -like "DotNet*"}}
    if ($UpdateGroup -eq 'Latest Cumulative Update LCU') {$AllOSDUpdates = $AllOSDUpdates | Where-Object {$_.UpdateGroup -eq 'LCU'}}
    if ($UpdateGroup -eq 'Servicing Stack Update SSU') {$AllOSDUpdates = $AllOSDUpdates | Where-Object {$_.UpdateGroup -eq 'SSU'}}
    #===================================================================================================
    #   GridView
    #===================================================================================================
    $AllOSDUpdates = $AllOSDUpdates | Select-Object -Property Catalog,OSDStatus,CreationDate,KBNumber,Title,FileName,Size,FileUri,OriginUri,OSDGuid
    if ($GridView.IsPresent) {$AllOSDUpdates = $AllOSDUpdates | Out-GridView -PassThru -Title "Select OSDUpdate Downloads"}
    #===================================================================================================
    #   Sort
    #===================================================================================================
    $AllOSDUpdates = $AllOSDUpdates | Sort-Object DateCreated
    #===================================================================================================
    #   Download
    #===================================================================================================
    if ($CatalogOffice) {
        foreach ($Update in $AllOSDUpdates) {
            $UpdateFile = $($Update.FileName)
            $MspFile = $UpdateFile -replace '.cab', '.msp'
            $DownloadDirectory = "$DownloadsPath\$($Update.Title)"

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
                Write-Host "Date Created: $($Update.DateCreated)" -ForegroundColor Gray
                Write-Host "Source: $DownloadDirectory\$MspFile" -ForegroundColor Gray
                Write-Host "Destination: $OfficeSetupUpdatesPath\$MspFile" -ForegroundColor Gray
                Copy-Item -Path "$DownloadDirectory\$MspFile" "$OfficeSetupUpdatesPath\$MspFile" -Force
                Write-Host ""
            }
        }
    }
    if ($CatalogWindows) {
        foreach ($Update in $AllOSDUpdates) {
            $UpdateFile = $($Update.FileName)
            $DownloadDirectory = "$DownloadsPath\$($Update.Title)"

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
    #   AddInstallScript
    #===================================================================================================
    if ($AddInstallScript.IsPresent) {
        Write-Verbose "Adding $RepositoryRootPath\Install-OSDUpdate.ps1" -Verbose
        Copy-Item "$($MyInvocation.MyCommand.Module.ModuleBase)\Scripts\Install-OSDUpdate.ps1" "$RepositoryRootPath" -Force | Out-Null
        if ($CatalogOffice) {
            Write-Verbose "Adding $DownloadsPath\OSDUpdate-Office.ps1" -Verbose
            Copy-Item "$($MyInvocation.MyCommand.Module.ModuleBase)\Scripts\OSDUpdate-Office.ps1" "$DownloadsPath" -Force | Out-Null
        }
        if ($CatalogWindows) {
            Write-Verbose "Adding $DownloadsPath\OSDUpdate-Windows.ps1" -Verbose
            Copy-Item "$($MyInvocation.MyCommand.Module.ModuleBase)\Scripts\OSDUpdate-Windows.ps1" "$DownloadsPath" -Force | Out-Null
        }
    }
    #===================================================================================================
}