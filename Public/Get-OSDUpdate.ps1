<#
.SYNOPSIS
Returns an Array of Microsoft Updates

.DESCRIPTION
Returns an Array of Microsoft Updates contained in the local WSUS Catalogs

.LINK
https://www.osdeploy.com/osdupdate/docs/functions/get-osdupdate

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
    if (!($Silent.IsPresent)) {Write-Warning "Updates are Current as of June 18, 2019"}
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
    $OSDUpdate = $OSDUpdate | Sort-Object -Property @{Expression = {$_.CreationDate}; Ascending = $false}, Size -Descending
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
