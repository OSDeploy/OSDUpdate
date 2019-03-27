<#
.SYNOPSIS
Creates an OSDUpdate Repository

.DESCRIPTION
Creates an OSDUpdate Repository to store downloaded updates

.LINK
https://www.osdeploy.com/osdupdate/docs/functions/new-osdupdaterepository

.PARAMETER Catalog
The Microsoft Office or Windows Update Catalog to store

.PARAMETER RepositoryRootPath
Full Path of the OSDUpdate Repository

.EXAMPLE
New-OSDUpdateRepository -Catalog 'Windows 10 x64' -RepositoryRootPath "C:\OSDUpdate"
Creates an OSDUpdate Repository in C:\OSDUpdate
Creates a directory C:\OSDUpdate\Windows 10 x64 with an initial Update Catalog
#>
function New-OSDUpdateRepository {
    [CmdletBinding()]
    PARAM (
        [Parameter(Mandatory = $True)]
        [ValidateSet(
            'Office 2010 32-Bit',
            'Office 2010 64-Bit',
            'Office 2013 32-Bit',
            'Office 2013 64-Bit',
            'Office 2016 32-Bit',
            'Office 2016 64-Bit',
            'Windows 7',
            #'Windows 8.1',
            #'Windows 8.1 Dynamic Update',
            'Windows 10',
            'Windows 10 Dynamic Update',
            'Windows 10 Feature On Demand',
            'Windows 10 Language Packs',
            'Windows 10 Language Interface Packs',
            #'Windows Server 2012 R2',
            'Windows Server 2016',
            'Windows Server 2019')]
            [string]$Catalog,
            
        [Parameter(Mandatory = $True)]
        [string]$RepositoryRootPath

    )
    if (!(Test-Path "$RepositoryRootPath")) {New-Item -Path "$RepositoryRootPath" -ItemType Directory -Force | Out-Null}
    if (!(Test-Path "$RepositoryRootPath\$Catalog")) {New-Item -Path "$RepositoryRootPath\$Catalog" -ItemType Directory -Force | Out-Null}
    Copy-Item -Path "$($MyInvocation.MyCommand.Module.ModuleBase)\Catalogs\$Catalog.xml" -Destination "$RepositoryRootPath\$Catalog" -Force | Out-Null
}