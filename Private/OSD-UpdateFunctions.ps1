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

function Get-MSPFileInfo {
    param
    (
        [Parameter(Mandatory = $true)][IO.FileInfo]$Path,
        [Parameter(Mandatory = $true)][ValidateSet('Classification', 'Description', 'DisplayName', 'KBArticle Number', 'ManufacturerName', 'ReleaseVersion', 'TargetProductName', 'Release', 'MoreInfoURL', 'OptimizedInstallMode', 'CreationTimeUTC', 'AllowRemoval', 'OptimizeCA', 'BuildNumber', 'StdPackageName', 'PatchType', 'IsMinorUpgrade')][string]$Property
    )
    
    try {
        #Creating windows installer object
        $WindowsInstaller = New-Object -ComObject WindowsInstaller.Installer
        
        #Loads the MSI database and specifies the mode to open it in by the last number on the line
        $MSIDatabase = $WindowsInstaller.GetType().InvokeMember("OpenDatabase", "InvokeMethod", $Null, $WindowsInstaller, @($Path.FullName, 32))
        
        #Specifies to query the MSIPatchMetadata table and get the value associated with the designated property
        $Query = "SELECT Value FROM MsiPatchMetadata WHERE Property = '$($Property)'"
        
        #Open up the property view
        $View = $MSIDatabase.GetType().InvokeMember("OpenView", "InvokeMethod", $null, $MSIDatabase, ($Query))
        $View.GetType().InvokeMember("Execute", "InvokeMethod", $null, $View, $null)
        
        #Retrieve the associate Property
        $Record = $View.GetType().InvokeMember("Fetch", "InvokeMethod", $null, $View, $null)
        
        #Retrieve the associated value of the retrieved property
        $Value = $Record.GetType().InvokeMember("StringData", "GetProperty", $null, $Record, 1)
        return $Value
        
    } catch {
        Write-Output $_.Exception.Message
    }
}