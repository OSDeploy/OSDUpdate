<#
.SYNOPSIS
Gets Current Microsoft Updates

.DESCRIPTION
Gets Current Microsoft Updates

.LINK
https://www.osdeploy.com/osdupdate/docs/functions/get-osdupdate

.PARAMETER CatalogOffice
The Microsoft Office Update Catalog selected for Updates

.PARAMETER OfficeProfile
Microsoft Office Update Type

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

.PARAMETER Silent
Hides the Warnings
#>
function Get-OSDUpdate {
    [CmdletBinding(DefaultParameterSetName = 'Office')]
    PARAM (
        #===================================================================================================
        #   Both
        #===================================================================================================
        [switch]$GridView,
        [switch]$Silent,
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
        [string]$UpdateGroup
    )
    #===================================================================================================
    #   Update Information
    #===================================================================================================
    if (!($Silent.IsPresent)) {Write-Warning "Updates are Current as of March 20, 2019"}
    #===================================================================================================
    #   Variables
    #===================================================================================================
    $AllOSDUpdates = @()
    #===================================================================================================
    #   Update Catalogs
    #===================================================================================================
    if ($CatalogOffice) {$UpdateCatalogs = Get-ChildItem -Path "$($MyInvocation.MyCommand.Module.ModuleBase)\Catalogs\*" -Include "$($CatalogOffice).xml"}
    if ($CatalogWindows) {$UpdateCatalogs = Get-ChildItem -Path "$($MyInvocation.MyCommand.Module.ModuleBase)\Catalogs\*" -Include "$($CatalogWindows).xml"}
    #===================================================================================================
    #   Import
    #===================================================================================================
    foreach ($UpdateCatalog in $UpdateCatalogs) {$AllOSDUpdates += Import-Clixml -Path "$($UpdateCatalog.FullName)"}
    #===================================================================================================
    #   Standard Filters
    #===================================================================================================
    $AllOSDUpdates = $AllOSDUpdates | Where-Object {$_.FileName -notlike "*.exe"}
    $AllOSDUpdates = $AllOSDUpdates | Where-Object {$_.FileName -notlike "*.psf"}
    $AllOSDUpdates = $AllOSDUpdates | Where-Object {$_.FileName -notlike "*.txt"}
    $AllOSDUpdates = $AllOSDUpdates | Where-Object {$_.FileName -notlike "*delta.exe"}
    $AllOSDUpdates = $AllOSDUpdates | Where-Object {$_.FileName -notlike "*express.cab"}
    #===================================================================================================
    #   Office Superseded
    #===================================================================================================
    if ($CatalogOffice) {
        $AllOSDUpdates = $AllOSDUpdates | Sort-Object OriginUri -Unique
        $AllOSDUpdates = $AllOSDUpdates | Sort-Object CreationDate -Descending

        $CurrentUpdates = @()
        $SupersededUpdates = @()

        foreach ($OfficeUpdate in $AllOSDUpdates) {
            $SkipUpdate = $false

            foreach ($CurrentUpdate in $CurrentUpdates) {
                if ($($OfficeUpdate.FileName) -eq $($CurrentUpdate.FileName)) {$SkipUpdate = $true}
            }

            if ($SkipUpdate) {
                if (!($Silent.IsPresent)) {
                    Write-Host "$($OfficeUpdate.KBNumber) $($OfficeUpdate.CreationDate) $($OfficeUpdate.Title)\$($OfficeUpdate.FileName) is superseded" -ForegroundColor DarkGray
                }
                $SupersededUpdates += $OfficeUpdate
                $SupersededUpdates | ConvertTo-Json | Out-File "$env:TEMP\$CatalogOffice.xml"
            } else {
                $CurrentUpdates += $OfficeUpdate
            }
        }
        $AllOSDUpdates = $CurrentUpdates
    }
    #===================================================================================================
    #   Office Profile
    #===================================================================================================
    if ($OfficeProfile -eq 'Default') {
        $AllOSDUpdates = $AllOSDUpdates | Where-Object {$_.FileName -like "*none*" -or $_.FileName -like "*en-us*"}
        $AllOSDUpdates = $AllOSDUpdates | Where-Object {$_.Title -notlike "*Language Pack*"}
    }

    if ($OfficeProfile -eq 'Language') {
        $AllOSDUpdates = $AllOSDUpdates | Where-Object {$_.FileName -notlike "*none*" -and $_.FileName -notlike "*en-us*"}
    }

    if ($OfficeProfile -eq 'Proofing') {
        $AllOSDUpdates = $AllOSDUpdates | Where-Object {$_.FileName -like "*Proof*"}
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
    $AllOSDUpdates = $AllOSDUpdates | Sort-Object -Property @{Expression = {$_.CreationDate}; Ascending = $false}, Size -Descending
    #===================================================================================================
    #   GridView
    #===================================================================================================
    if ($GridView.IsPresent) {$AllOSDUpdates = $AllOSDUpdates | Out-GridView -PassThru -Title 'Select OSDUpdates'}
    #===================================================================================================
    #   Return
    #===================================================================================================
    Return $AllOSDUpdates
}
