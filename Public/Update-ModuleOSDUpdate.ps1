<#
.SYNOPSIS
Updates the OSDUpdate PowerShell Module to the latest version

.DESCRIPTION
Updates the OSDUpdate PowerShell Module to the latest version from the PowerShell Gallery

.LINK
https://www.osdeploy.com/osdupdate/docs/functions/update-moduleosdupdate

.Example
Update-ModuleOSDUpdate
#>
function Update-ModuleOSDUpdate {
    [CmdletBinding()]
    PARAM ()
    try {
        Write-Warning "Uninstall-Module -Name OSDUpdate -AllVersions -Force"
        Uninstall-Module -Name OSDUpdate -AllVersions -Force
    }
    catch {}

    try {
        Write-Warning "Install-Module -Name OSDUpdate -Force"
        Install-Module -Name OSDUpdate -Force
    }
    catch {}

    try {
        Write-Warning "Import-Module -Name OSDUpdate -Force"
        Import-Module -Name OSDUpdate -Force
    }
    catch {}
}