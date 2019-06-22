#Requires -Modules OSDUpdate
#Requires -Version 5

<#
.SYNOPSIS
    Uses the OSDUpdate Module to update subdirectory OSDUpdate Packages

.DESCRIPTION
    Uses the OSDUpdate Module to update subdirectory OSDUpdate Packages

.NOTES
    Author:         David Segura
    Website:        osdeploy.com
    Twitter:        @SeguraOSD
    Version:        19.6.21.0
#>

#======================================================================================
#   Set Error Action to Silently Continue
#======================================================================================
$ErrorActionPreference = "SilentlyContinue"

#======================================================================================
#   Import OSDUpdate Module
#======================================================================================
Import-Module OSDUpdate -Force

#======================================================================================
#   Get all Package Subdirectories
#======================================================================================
$Packages = Get-ChildItem $PSScriptRoot -Directory | Select-Object -Property Name

#======================================================================================
#   Process Package Directories
#   Only directories named with a proper PackageName will be updated
#======================================================================================
foreach ($Package in $Packages) {
    Write-Host "Updating Package $($Package.Name)" -ForegroundColor Cyan
    New-OSDUpdatePackage -PackageName "$($Package.Name)" -PackagePath $PSScriptRoot -AppendPackageName -RemoveSuperseded
}