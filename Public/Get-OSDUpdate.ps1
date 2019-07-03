<#
.SYNOPSIS
Returns an Array of Microsoft Updates

.DESCRIPTION
Returns an Array of Microsoft Updates contained in the local WSUS Catalogs

.LINK
https://www.osdeploy.com/osdupdate/functions/get-osdupdate

.PARAMETER GridView
Displays the results in GridView with -PassThru

.PARAMETER Silent
Hide the Current Update Date information
#>

function Get-OSDUpdate {
    [CmdletBinding()]
    PARAM (
        [switch]$GridView,
        [switch]$Silent
    )
    #===================================================================================================
    #   Update Information
    #===================================================================================================
    if (!($Silent.IsPresent)) {
        Write-Verbose 'Updates are Current as of July 2, 2019' -Verbose
        Write-Verbose 'Gathering Updates ... Please Wait ...' -Verbose

    }
    #===================================================================================================
    #   Variables
    #===================================================================================================
    $OSDUpdate = @()
    #===================================================================================================
    #   UpdateCatalogs
    #===================================================================================================
    $OSDUpdateCatalogs = Get-ChildItem -Path "$($MyInvocation.MyCommand.Module.ModuleBase)\Catalogs\*" -Include "*.xml"
    #===================================================================================================
    #   Import Catalog XML Files
    #===================================================================================================
    foreach ($OSDUpdateCatalog in $OSDUpdateCatalogs) {
<#         #Write-Verbose "Importing $($OSDUpdateCatalog.Name)" -Verbose
        if ($OSDUpdateCatalog.Name -match 'Office') {

            $OfficeUpdates = @()
            $OfficeUpdates = Import-Clixml -Path "$($OSDUpdateCatalog.FullName)"

            $OfficeUpdates = $OfficeUpdates | Sort-Object OriginUri -Unique
            $OfficeUpdates = $OfficeUpdates | Sort-Object CreationDate -Descending

            #$OfficeUpdates | Out-GridView

            $CurrentUpdates = @()
            $SupersededUpdates = @()

            foreach ($OfficeUpdate in $OfficeUpdates) {
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
            $OSDUpdate += $CurrentUpdates
        } else {
            $OSDUpdate += Import-Clixml -Path "$($OSDUpdateCatalog.FullName)"
        } #>
        $OSDUpdate += Import-Clixml -Path "$($OSDUpdateCatalog.FullName)"
    }
    #===================================================================================================
    #   Standard Filters
    #===================================================================================================
    $OSDUpdate = $OSDUpdate | Where-Object {$_.FileName -notlike "*.exe"}
    $OSDUpdate = $OSDUpdate | Where-Object {$_.FileName -notlike "*.psf"}
    $OSDUpdate = $OSDUpdate | Where-Object {$_.FileName -notlike "*.txt"}
    $OSDUpdate = $OSDUpdate | Where-Object {$_.FileName -notlike "*delta.exe"}
    $OSDUpdate = $OSDUpdate | Where-Object {$_.FileName -notlike "*express.cab"}
    #===================================================================================================
    #   Sorting
    #===================================================================================================
    #$OSDUpdate = $OSDUpdate | Sort-Object -Property @{Expression = {$_.CreationDate}; Ascending = $false}, Size -Descending
    $OSDUpdate = $OSDUpdate | Sort-Object -Property CreationDate -Descending
    #===================================================================================================
    #   GridView
    #===================================================================================================
    if ($GridView.IsPresent) {
        $OSDUpdate = $OSDUpdate | Out-GridView -PassThru -Title 'Select OSDUpdates to Return'
    }
    #===================================================================================================
    #   Return
    #===================================================================================================
    Return $OSDUpdate
}
