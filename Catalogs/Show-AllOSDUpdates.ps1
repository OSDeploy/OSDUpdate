$AllOSDUpdates = @()
$AllUpdateCatalogs = Get-ChildItem -Path "$PSScriptRoot\*" -Include '*.xml'
foreach ($UpdateCatalog in $AllUpdateCatalogs) {$AllOSDUpdates += Import-Clixml -Path "$($UpdateCatalog.FullName)"}

$AllOSDUpdates = $AllOSDUpdates | Select-Object -Property * | Sort-Object CreationDate -Descending | Out-GridView -PassThru -Title "All OSDUpdates"
Write-Host ""
$AllOSDUpdates = $AllOSDUpdates | Select-Object -Property Title
$AllOSDUpdates | Sort-Object Title -Unique | Format-Table

Write-Host ""
Pause