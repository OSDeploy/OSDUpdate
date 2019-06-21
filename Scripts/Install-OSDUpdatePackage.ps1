#Requires -RunAsAdministrator
#Requires -Version 5

<#
.SYNOPSIS
    Installs Updates in an OSDUpdate Package

.DESCRIPTION
    Installs Updates in an OSDUpdate Package

.NOTES
    Author:         David Segura
    Website:        osdeploy.com
    Twitter:        @SeguraOSD
    Version:        19.6.21.0
#>

function Convert-GuidToCompressedGuid {
    <#
    .SYNOPSIS
        This converts a GUID to a compressed GUID also known as a product code.	
    .DESCRIPTION
        This function will typically be used to figure out the product code
        that matches up with the product code stored in the 'SOFTWARE\Classes\Installer\Products'
        registry path to a MSI installer GUID.
    .EXAMPLE
        Convert-GuidToCompressedGuid -Guid '{7C6F0282-3DCD-4A80-95AC-BB298E821C44}'
    
        This example would output the compressed GUID '2820F6C7DCD308A459CABB92E828C144'
    .PARAMETER Guid
        The GUID you'd like to convert.
    .LINK
        https://www.adamtheautomator.com/compressed-guid-with-powershell/
    #>
    [CmdletBinding()]
    [OutputType()]
    param (
        [Parameter(ValueFromPipeline, ValueFromPipelineByPropertyName, Mandatory)]
        [string]$Guid
    )
    begin {
        $Guid = $Guid.Replace('-', '').Replace('{', '').Replace('}', '')
    }
    process {
        try {
            $Groups = @(
                $Guid.Substring(0, 8).ToCharArray(),
                $Guid.Substring(8, 4).ToCharArray(),
                $Guid.Substring(12, 4).ToCharArray(),
                $Guid.Substring(16, 16).ToCharArray()
            )
            $Groups[0..2] | foreach {
                [array]::Reverse($_)
            }
            $CompressedGuid = ($Groups[0..2] | foreach { $_ -join '' }) -join ''
            
            $chararr = $Groups[3]
            for ($i = 0; $i -lt $chararr.count; $i++) {
                if (($i % 2) -eq 0) {
                    $CompressedGuid += ($chararr[$i+1] + $chararr[$i]) -join ''
                }
            }
            $CompressedGuid
        } catch {
            Write-Error $_.Exception.Message	
        }
    }
}

function Convert-CompressedGuidToGuid {
<#
    .SYNOPSIS
        This converts a compressed GUID also known as a product code into a GUID.	
    .DESCRIPTION
        This function will typically be used to figure out the MSI installer GUID
        that matches up with the product code stored in the 'SOFTWARE\Classes\Installer\Products'
        registry path.
    .EXAMPLE
        Convert-CompressedGuidToGuid -CompressedGuid '2820F6C7DCD308A459CABB92E828C144'
    
        This example would output the GUID '{7C6F0282-3DCD-4A80-95AC-BB298E821C44}'
    .PARAMETER CompressedGuid
        The compressed GUID you'd like to convert.
    .LINK
        https://www.adamtheautomator.com/convert-compressed-guid-to-guid/
    #>
    [CmdletBinding()]
    [OutputType([System.String])]
    param (
        [Parameter(ValueFromPipeline, ValueFromPipelineByPropertyName, Mandatory)]
        [ValidatePattern('^[0-9a-fA-F]{32}$')]
        [string]$CompressedGuid
    )
    process {
        $Indexes = [ordered]@{
            0 = 8;
            8 = 4;
            12 = 4;
            16 = 2;
            18 = 2;
            20 = 2;
            22 = 2;
            24 = 2;
            26 = 2;
            28 = 2;
            30 = 2
        }
        #$Guid = '{'
        $Guid = ''
        foreach ($index in $Indexes.GetEnumerator()) {
            $part = $CompressedGuid.Substring($index.Key, $index.Value).ToCharArray()
            [array]::Reverse($part)
            $Guid += $part -join ''
        }
        $Guid = $Guid.Insert(9,'-').Insert(14, '-').Insert(19, '-').Insert(24, '-')
        #$Guid + '}'
        $Guid + ''
    }
}
#======================================================================================
#   Validate Admin Rights
#======================================================================================
Write-Host ""
# Verify Running as Admin
$isAdmin = ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")
If (!( $isAdmin )) {
    Write-Host "Checking User Account Control settings ..." -ForegroundColor Green
    if ((Get-ItemProperty HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System).EnableLUA -eq 0) {
        #UAC Disabled
        Write-Host '========================================================================================' -ForegroundColor DarkGray
        Write-Host "User Account Control is Disabled ... " -ForegroundColor Green
        Write-Host "You will need to correct your UAC Settings ..." -ForegroundColor Green
        Write-Host "Try running this script in an Elevated PowerShell session ... Exiting" -ForegroundColor Green
        Write-Host '========================================================================================' -ForegroundColor DarkGray
        Start-Sleep -s 10
        Exit 0
    } else {
        #UAC Enabled
        Write-Host "UAC is Enabled" -ForegroundColor Green
        Start-Sleep -s 3
        if ($Silent) {
            Write-Host "-- Restarting as Administrator (Silent)" -ForegroundColor Cyan ; Start-Sleep -Seconds 1
            Start-Process powershell.exe "-NoProfile -ExecutionPolicy Bypass -File `"$PSCommandPath`" -Silent" -Verb RunAs -Wait
        } elseif($Restart) {
            Write-Host "-- Restarting as Administrator (Restart)" -ForegroundColor Cyan ; Start-Sleep -Seconds 1
            Start-Process powershell.exe "-NoProfile -ExecutionPolicy Bypass -File `"$PSCommandPath`" -Restart" -Verb RunAs -Wait
        } else {
            Write-Host "-- Restarting as Administrator" -ForegroundColor Cyan ; Start-Sleep -Seconds 1
            Start-Process powershell.exe "-NoProfile -ExecutionPolicy Bypass -File `"$PSCommandPath`"" -Verb RunAs -Wait
        }
        Exit 0
    }
} else {
    Write-Host '========================================================================================' -ForegroundColor DarkGray
    Write-Host "-- Running with Elevated Permissions ..." -ForegroundColor Cyan ; Start-Sleep -Seconds 1
    Write-Host '========================================================================================' -ForegroundColor DarkGray
}
#======================================================================================
#   Begin
#======================================================================================
Write-Host "Install-OSDUpdate Package" -ForegroundColor Green
#======================================================================================
#   OS Information
#======================================================================================
$OSCaption = $((Get-WmiObject -Class Win32_OperatingSystem).Caption).Trim()
$OSArchitecture = $((Get-WmiObject -Class Win32_OperatingSystem).OSArchitecture).Trim()
$OSProductType = $((Get-WmiObject -Class Win32_OperatingSystem).ProductType)
$OSVersion = $((Get-WmiObject -Class Win32_OperatingSystem).Version).Trim()
$OSBuildNumber = $((Get-WmiObject -Class Win32_OperatingSystem).BuildNumber).Trim()
Write-Host "Operating System: $OSCaption" -ForegroundColor Cyan
Write-Host "OS Architecture: $OSArchitecture" -ForegroundColor Cyan
Write-Host "OS ProductType: $OSProductType" -ForegroundColor Cyan
Write-Host "OS Version: $OSVersion" -ForegroundColor Cyan
Write-Host "OS Build Number: $OSBuildNumber" -ForegroundColor Cyan
if ($OSVersion -Like "10*") {
    $OSReleaseID = (Get-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion" -Name ReleaseId).ReleaseId
    Write-Host "OS Release ID: $OSReleaseID" -ForegroundColor Cyan
}
#======================================================================================
#   Script Information
#======================================================================================
$Invocation = (Get-Variable MyInvocation -Scope Script).Value
$ScriptPath = Split-Path -Parent $Invocation.MyCommand.Path
#======================================================================================
#   Get OSDUpdate Package Catalog
#======================================================================================
$Updates = @()
$UpdateCatalogs = Get-ChildItem $PSScriptRoot "OSDUpdatePackage.xml"
Try {
    foreach ($Catalog in $UpdateCatalogs) {
        $Updates += Import-Clixml -Path $Catalog.FullName
    }
}
Catch {
    Write-Warning -Message "Could not locate OSDUpdatePackage.xml ... Exiting"
    Break
}
#======================================================================================
#   Determine Catalog
#======================================================================================
$Catalog = ($Updates | Select-Object -First 1).Catalog
Write-Host "OSDUpdate Catalog: $Catalog" -ForegroundColor Cyan
#======================================================================================
#   Catalog Office
#======================================================================================
if ($Catalog -like "Office*") {
    #======================================================================================
    #   Installed Patches
    #======================================================================================
    $PatchesInstalledRegistry = @()
    $PatchesInstalledRegistry = 'HKLM:\SOFTWARE\Classes\Installer\Patches'
    $PatchesInstalledProductCode = @()
    $PatchesInstalledProductCode = Get-ChildItem -Path $PatchesInstalledRegistry -EA SilentlyContinue | Select-Object -Property @{Name="ProductCode"; Expression = {$_.PSChildName}} -Unique
    $PatchesInstalledGuids = @()
    foreach ($InstalledPatch in $PatchesInstalledProductCode) {
        $InstalledPatchGuid = Convert-CompressedGuidToGuid -CompressedGuid "$($InstalledPatch.ProductCode)"
        $PatchesInstalledGuids += $InstalledPatchGuid
    }
    #======================================================================================
    #   Available Patches (MSP's)
    #======================================================================================
    $PatchesAvailable = @()
    $PatchesAvailable = Get-ChildItem "$ScriptPath" -Recurse -File -Include *.msp | Select-Object -Property LastWriteTime,Name,Length,FullName,Directory,BaseName,Extension
    $PatchesAvailable = $PatchesAvailable | Sort-Object -Property @{Expression = {$_.LastWriteTime}; Ascending = $true}, Length -Descending
    #======================================================================================
    #   Get Patch XML Information
    #======================================================================================
    foreach ($Patch in $PatchesAvailable) {
        $PatchXml = "$($Patch.Directory)\$($Patch.BaseName).xml"
    
        $Patch | Add-Member -MemberType NoteProperty -Name PatchGuid -value ''
        $Patch | Add-Member -MemberType NoteProperty -Name ProductCode -value ''
        $Patch | Add-Member -MemberType NoteProperty -Name TargetProductCode -value ''
    
        if (Test-Path $PatchXml) {
            $xml = [xml](Get-Content $PatchXml)
            $Patch.PatchGuid = $($xml.MsiPatch | Select PatchGuid).PatchGuid
            $Patch.ProductCode = Convert-GuidToCompressedGuid -Guid $($Patch.PatchGuid)
            $Patch.TargetProductCode = $($xml.MsiPatch.TargetProductCode)
        }
    }
    #======================================================================================
    #   Set InstallationStatus
    #======================================================================================
    foreach ($Patch in $PatchesAvailable) {
        $Patch | Add-Member -MemberType NoteProperty -Name InstallStatus -value ''
        foreach ($PatchInstalled in $PatchesInstalledProductCode) {
            if ($Patch.ProductCode -eq $PatchInstalled.ProductCode) {
                $Patch.InstallStatus = 'Installed'
            }
        }
    }
    #======================================================================================
    #   Install Updates
    #======================================================================================
    foreach ($Patch in $PatchesAvailable) {
        $PatchName = $($Patch.Directory) | Split-Path -Leaf
    
        if ($Patch.InstallStatus -eq 'Installed') {
            Write-Host "Installed: $PatchName $($Patch.Name)" -ForegroundColor DarkGray
        } else {
            Write-Host "$PatchName $($Patch.Name)" -ForegroundColor Cyan
            msiexec /p "$($Patch.FullName)" /qn REBOOT=ReallySuppress MSIRESTARTMANAGERCONTROL=Disable | Out-Null
        }
    }
}
#======================================================================================
#   Catalog Windows
#======================================================================================
if ($Catalog -like "Windows*") {
    #======================================================================================
    #   Sessions
    #======================================================================================
    [xml]$SessionsXML = Get-Content -Path "$env:WinDir\Servicing\Sessions\Sessions.xml"

    $Sessions = $SessionsXML.SelectNodes('Sessions/Session') | ForEach-Object {
        New-Object -Type PSObject -Property @{
            Id = $_.Tasks.Phase.package.id
            KBNumber = $_.Tasks.Phase.package.name
            TargetState = $_.Tasks.Phase.package.targetState
            Client = $_.Client
            Complete = $_.Complete
            Status = $_.Status
        }
    }
    $Sessions = $Sessions | Where-Object {$_.Id -like "Package*"}
    $Sessions = $Sessions | Select-Object -Property Id, KBNumber, TargetState, Client, Status, Complete | Sort-Object Complete -Descending
    #======================================================================================
    #   Architecture
    #======================================================================================
    if ($OSArchitecture -like "*64*") {$Updates = $Updates | Where-Object {$_.UpdateArch -eq 'x64'}}
    else {$Updates = $Updates | Where-Object {$_.UpdateArch -eq 'x86'}}
    #======================================================================================
    #   OSReleaseID
    #======================================================================================
    if ($OSProductType -eq 1) {
        $Updates = $Updates | Where-Object {$_.Catalog -notlike "*Server*"}
        if ($OSVersion -like "6.1") {$Updates = $Updates | Where-Object {$_.Catalog -like "Windows 7*"}}
    } else {
        $Updates = $Updates | Where-Object {$_.Catalog -like "*Server*"}
    }
    if ($OSVersion -like "10.*") {$Updates = $Updates | Where-Object {$_.UpdateBuild -eq $OSReleaseID}}
    #======================================================================================
    #   Get-Hotfix
    #======================================================================================
    $InstalledUpdates = Get-HotFix
    #======================================================================================
    #   Windows Updates
    #======================================================================================
    Write-Host "Updating Windows" -ForegroundColor Green
    foreach ($Update in $Updates) {
        if ($Update.UpdateGroup -eq 'SSU') {
            $UpdatePath = "$PSScriptRoot\$($Update.Title)\$($Update.FileName)"
            if (Test-Path "$UpdatePath") {
                Write-Host "$UpdatePath" -ForegroundColor DarkGray
                if ($InstalledUpdates | Where-Object HotFixID -like "*$($Update.KBNumber)") {
                    Write-Host "KB$($Update.KBNumber) is already installed" -ForegroundColor Cyan
                } else {
                    Write-Host "Installing $($Update.Title) ..." -ForegroundColor Cyan
                    Dism /Online /Add-Package /PackagePath:"$UpdatePath" /NoRestart
                }
            } else {
                Write-Warning "Not Found: $UpdatePath"
            }
        }
    }
    foreach ($Update in $Updates) {
        if ($Update.UpdateGroup -ne 'SSU') {
            $UpdatePath = "$PSScriptRoot\$($Update.Title)\$($Update.FileName)"
            if (Test-Path "$UpdatePath") {
                Write-Host "$UpdatePath" -ForegroundColor DarkGray
                if ($InstalledUpdates | Where-Object HotFixID -like "*$($Update.KBNumber)") {
                    Write-Host "KB$($Update.KBNumber) is already installed" -ForegroundColor Cyan
                } else {
                    Write-Host "Installing $($Update.Title) ..." -ForegroundColor Cyan
                    Dism /Online /Add-Package /PackagePath:"$UpdatePath" /NoRestart
                }
            } else {
                #Write-Warning "Not Found: $UpdatePath"
            }
        }
    }
}