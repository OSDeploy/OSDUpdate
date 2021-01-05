#Requires -Version 5

<#
.SYNOPSIS
    Installs Updates in an OSDUpdate Package from Child Directories

.DESCRIPTION
    Installs Updates in an OSDUpdate Package from Child Directories

.NOTES
    Author:         David Segura
    Website:        osdeploy.com
    Twitter:        @SeguraOSD
    Version:        20.1.5.1
#>
#======================================================================================
#   Begin
#======================================================================================
Write-Host "Installing Windows Defender Updates" -ForegroundColor Green
#======================================================================================
#   Execute
#======================================================================================
Get-ChildItem -Path $PSScriptRoot mpam*.exe | foreach {
    Write-Host $_.FullName -ForegroundColor Cyan
    Start-Process $_.FullName -Wait
}
#======================================================================================
#   Complete
#======================================================================================