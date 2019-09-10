<# if (!$(Get-Module -ListAvailable OSDSUS)) {
    try {
        Install-Module OSDSUS -ErrorAction Stop
    }
    catch {
        Write-Error $_
        Write-Error "Problem installing Module A dependency Module OSDSUS! Module A will NOT be loaded. Halting!"
        $global:FunctionResult = "1"
        return
    }
}
try {
    Import-Module OSDSUS -ErrorAction Stop
}
catch {
    Write-Error $_
    Write-Error "Problem importing Module A dependency Module OSDSUS! Module A will NOT be loaded. Halting!"
    $global:FunctionResult = "1"
    return
} #>

#===================================================================================================
#   Import Functions
#   https://github.com/RamblingCookieMonster/PSStackExchange/blob/master/PSStackExchange/PSStackExchange.psm1
#===================================================================================================
$OSDPublicFunctions  = @( Get-ChildItem -Path $PSScriptRoot\Public\*.ps1 -ErrorAction SilentlyContinue )
$OSDPrivateFunctions = @( Get-ChildItem -Path $PSScriptRoot\Private\*.ps1 -ErrorAction SilentlyContinue )

foreach ($Import in @($OSDPublicFunctions + $OSDPrivateFunctions)) {
    Try {. $Import.FullName}
    Catch {Write-Error -Message "Failed to import function $($Import.FullName): $_"}
}

Export-ModuleMember -Function $OSDPublicFunctions.BaseName
#===================================================================================================