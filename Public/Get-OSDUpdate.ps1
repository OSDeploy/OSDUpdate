<#
.SYNOPSIS
Returns an Array of Microsoft Updates

.DESCRIPTION
Returns an Array of Microsoft Updates contained in the local WSUS Catalogs

.LINK
https://osdupdate.osdeploy.com/module/functions/get-osdupdate

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
    #   Variables
    #===================================================================================================
    $OSDUpdate = @()
    $OSDUpdate = Get-OSDSUS
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
