#Requires -Version 5

<#
.SYNOPSIS
    Standalone script for applying Legacy Office Updates

.DESCRIPTION
    Standalone script for applying Legacy Office Updates

.NOTES
    Author:         David Segura
    Website:        osdeploy.com
    Twitter:        @SeguraOSD
    Version:        19.6.25.0
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
#   Begin
#======================================================================================
Write-Host "OSDUpdate Microsoft Office" -ForegroundColor Green
#======================================================================================
#   Current Path
#======================================================================================
$Invocation = (Get-Variable MyInvocation -Scope Script).Value
$ScriptPath = Split-Path -Parent $Invocation.MyCommand.Path
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
