#https://github.com/RamblingCookieMonster/PSStackExchange/blob/master/PSStackExchange/PSStackExchange.psm1
$PublicFunctions  = @( Get-ChildItem -Path $PSScriptRoot\Public\*.ps1 -ErrorAction SilentlyContinue )
$PrivateFunctions = @( Get-ChildItem -Path $PSScriptRoot\Private\*.ps1 -ErrorAction SilentlyContinue )

foreach ($Import in @($PublicFunctions + $PrivateFunctions)) {
    Try {. $Import.FullName}
    Catch {Write-Error -Message "Failed to import function $($Import.FullName): $_"}
}

Export-ModuleMember -Function $PublicFunctions.BaseName