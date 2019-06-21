function New-OSDUpdatePackage {
    [CmdletBinding()]
    PARAM (
        [Parameter(Mandatory = $True)]
        [string]$PackagePath,
        [switch]$AppendPackageName,
        [switch]$RemoveSuperseded,
        [switch]$GridView,

        [Parameter(Mandatory = $True)]
        [ValidateSet(
            'Office 2010 32-Bit',
            'Office 2010 64-Bit',
            'Office 2013 32-Bit',
            'Office 2013 64-Bit',
            'Office 2016 32-Bit',
            'Office 2016 64-Bit',
            'Windows 7 x64',
            'Windows 7 x86',

            'Windows 10 x64 1803',
            'Windows 10 x64 1809',
            'Windows 10 x64 1903',

            'Windows 10 x86 1803',
            'Windows 10 x86 1809',
            'Windows 10 x86 1903',

            'Windows Server 2016 1607',
            'Windows Server 2016 1709',
            'Windows Server 2016 1803',

            'Windows Server 2019 1809')]
        [string]$PackageName,

        [ValidateSet('Default','Proofing','Language','All')]
        [string]$OfficeProfile = 'Default',

        [Alias('OfficeSetupUpdatesPath')]
        [string]$OfficeMediaUpdatesPath
    )

    #===================================================================================================
    #   Get-OSDUpdate
    #===================================================================================================
    $OSDUpdate = @()
    $OSDUpdate = Get-OSDUpdate
    #===================================================================================================
    #   AppendPackageName
    #===================================================================================================
    if ($AppendPackageName) {
        $PackagePath = "$PackagePath\$PackageName"
    }
    #===================================================================================================
    #   PackagePath
    #===================================================================================================
    if (!(Test-Path "$PackagePath")) {New-Item -Path "$PackagePath" -ItemType Directory -Force | Out-Null}
    #===================================================================================================
    #   Multi PackageName
    #===================================================================================================
    if ($PackageName -like "Office*") {
        $OSDUpdate = $OSDUpdate | Where-Object {$_.Catalog -eq $PackageName}
    }

    if ($PackageName -like "Windows*") {
        if ($PackageName -like "Windows 7*") {
            $OSDUpdate = $OSDUpdate | Where-Object {$_.Catalog -eq 'Windows 7'}
        }
        if ($PackageName -like "Windows 10*") {
            $OSDUpdate = $OSDUpdate | Where-Object {$_.Catalog -eq 'Windows 10'}
        }
        if ($PackageName -like "Windows Server 2016*") {
            $OSDUpdate = $OSDUpdate | Where-Object {$_.Catalog -eq 'Windows Server 2016'}
        }
        if ($PackageName -like "Windows Server 2019*") {
            $OSDUpdate = $OSDUpdate | Where-Object {$_.Catalog -eq 'Windows Server 2019'}
        }
    }
    #===================================================================================================
    #   Multi Filter
    #===================================================================================================
    if ($PackageName -like "Windows*") {
        if ($PackageName -like "*x64*") {$OSDUpdate = $OSDUpdate | Where-Object {$_.UpdateArch -eq 'x64'}}
        if ($PackageName -like "*x86*") {$OSDUpdate = $OSDUpdate | Where-Object {$_.UpdateArch -eq 'x86'}}
        if ($PackageName -like "*1507*") {$OSDUpdate = $OSDUpdate | Where-Object {$_.UpdateBuild -eq '1507'}}
        if ($PackageName -like "*1511*") {$OSDUpdate = $OSDUpdate | Where-Object {$_.UpdateBuild -eq '1511'}}
        if ($PackageName -like "*1607*") {$OSDUpdate = $OSDUpdate | Where-Object {$_.UpdateBuild -eq '1607'}}
        if ($PackageName -like "*1703*") {$OSDUpdate = $OSDUpdate | Where-Object {$_.UpdateBuild -eq '1703'}}
        if ($PackageName -like "*1709*") {$OSDUpdate = $OSDUpdate | Where-Object {$_.UpdateBuild -eq '1709'}}
        if ($PackageName -like "*1803*") {$OSDUpdate = $OSDUpdate | Where-Object {$_.UpdateBuild -eq '1803'}}
        if ($PackageName -like "*1809*") {$OSDUpdate = $OSDUpdate | Where-Object {$_.UpdateBuild -eq '1809'}}
        if ($PackageName -like "*1903*") {$OSDUpdate = $OSDUpdate | Where-Object {$_.UpdateBuild -eq '1903'}}
    }
    $OSDUpdate | Export-Clixml -Path "$PackagePath\OSDUpdatePackage.xml" -Force | Out-Null

    if ($PackageName -like "Office*") {
        if ($OfficeProfile -eq 'Default') {
            $OSDUpdate = $OSDUpdate | Where-Object {$_.FileName -like "*none*" -or $_.FileName -like "*en-us*"}
            $OSDUpdate = $OSDUpdate | Where-Object {$_.Title -notlike "*Language Pack*"}
        }
        if ($OfficeProfile -eq 'Language') {
            $OSDUpdate = $OSDUpdate | Where-Object {$_.FileName -notlike "*none*" -and $_.FileName -notlike "*en-us*"}
        }
        if ($OfficeProfile -eq 'Proofing') {
            $OSDUpdate = $OSDUpdate | Where-Object {$_.FileName -like "*Proof*"}
        }
    }
    #===================================================================================================
    #   Multi Existing Updates
    #===================================================================================================
    $ExistingUpdates = @()
    $SupersededUpdates = @()

    $ExistingUpdates = Get-ChildItem -Path "$PackagePath\*" -Directory -Recurse | Select-Object -Property *

    foreach ($Update in $ExistingUpdates) {
        if ($OSDUpdate.Title -NotContains $Update.Name) {$SupersededUpdates += $Update.FullName}
    }
    #===================================================================================================
    #   Multi Superseded Updates
    #===================================================================================================
    foreach ($Update in $SupersededUpdates) {
        if ($RemoveSuperseded.IsPresent) {
            Write-Warning "Removing Superseded: $Update"
            Remove-Item $Update -Recurse -Force | Out-Null
        } else {
            Write-Warning "Superseded: $Update"
        }
    }
    #===================================================================================================
    #   Multi Get Downloaded Updates
    #===================================================================================================
    foreach ($Update in $OSDUpdate) {
        if ($PackageName -like "Windows*") {
            $FullUpdatePath = "$PackagePath\$($Update.Title)\$($Update.FileName)"
            if (Test-Path $FullUpdatePath) {
                $Update.OSDStatus = "Downloaded"
            }
        }
    }
    #===================================================================================================
    #   Multi GridView
    #===================================================================================================
    if ($PackageName -like "Office*") {
        $OSDUpdate = $OSDUpdate | Select-Object -Property OSDStatus,Catalog,CreationDate,KBNumber,Title,FileName,Size,FileUri,OriginUri,OSDGuid
    }
    if ($PackageName -like "Windows*") {
        $OSDUpdate = $OSDUpdate | Select-Object -Property OSDStatus,Catalog,UpdateOS,UpdateArch,UpdateBuild,CreationDate,KBNumber,Title,FileName,Size,FileUri,OriginUri,OSDGuid
    }
    if ($GridView.IsPresent) {$OSDUpdate = $OSDUpdate | Out-GridView -PassThru -Title "Select OSDUpdate Downloads to include in the Package"}
    #===================================================================================================
    #   Sort
    #===================================================================================================
    $OSDUpdate = $OSDUpdate | Sort-Object DateCreated
    #===================================================================================================
    #   Office Download
    #===================================================================================================
    if ($PackageName -like "Office*") {
        foreach ($Update in $OSDUpdate) {
            $UpdateFile = $($Update.FileName)
            $MspFile = $UpdateFile -replace '.cab', '.msp'
            $DownloadDirectory = "$PackagePath\$($Update.Title)"

            if (!(Test-Path "$DownloadDirectory")) {New-Item -Path "$DownloadDirectory" -ItemType Directory -Force | Out-Null}
        
            if (Test-Path "$DownloadDirectory\$MspFile") {
                Write-Host "$DownloadDirectory\$MspFile" -ForegroundColor Cyan
            } else {
                Write-Host "$DownloadDirectory\$MspFile" -ForegroundColor Cyan
                Write-Host "Download: $($Update.OriginUri)" -ForegroundColor Gray
                Start-BitsTransfer -Source $($Update.OriginUri) -Destination "$DownloadDirectory\$UpdateFile"
            }

            if ((Test-Path "$DownloadDirectory\$UpdateFile") -and (!(Test-Path "$DownloadDirectory\$MspFile"))) {
                Write-Host "Expand: $DownloadDirectory\$MspFile" -ForegroundColor Gray
                expand "$DownloadDirectory\$UpdateFile" -F:* "$DownloadDirectory" | Out-Null
            }

            if ((Test-Path "$DownloadDirectory\$UpdateFile") -and (Test-Path "$DownloadDirectory\$MspFile")) {
                Write-Host "Remove: $DownloadDirectory\$UpdateFile" -ForegroundColor Gray
                Remove-Item "$DownloadDirectory\$UpdateFile" -Force | Out-Null
            }
        }
        #===================================================================================================
        #   Office Setup Updates
        #===================================================================================================
        if ($OfficeMediaUpdatesPath) {
            if (!(Test-Path "$OfficeMediaUpdatesPath")) {New-Item -Path "$OfficeMediaUpdatesPath" -ItemType Directory -Force | Out-Null}
            Write-Host "Date Created: $($Update.DateCreated)" -ForegroundColor Gray
            Write-Host "Source: $DownloadDirectory\$MspFile" -ForegroundColor Gray
            Write-Host "Destination: $OfficeMediaUpdatesPath\$MspFile" -ForegroundColor Gray
            Copy-Item -Path "$DownloadDirectory\$MspFile" "$OfficeMediaUpdatesPath\$MspFile" -Force
            Write-Host ""
        }
    }
    #===================================================================================================
    #   Windows Download
    #===================================================================================================
    if ($PackageName -like "Windows*") {
        foreach ($Update in $OSDUpdate) {
            $UpdateFile = $($Update.FileName)
            $DownloadDirectory = "$PackagePath\$($Update.Title)"

            if (!(Test-Path "$DownloadDirectory")) {New-Item -Path "$DownloadDirectory" -ItemType Directory -Force | Out-Null}
        
            if (Test-Path "$DownloadDirectory\$UpdateFile") {
                Write-Host "$($Update.Title)" -ForegroundColor Cyan
                Write-Host "$DownloadDirectory\$UpdateFile" -ForegroundColor Gray
                #Write-Host "Update already downloaded" -ForegroundColor Gray
            } else {
                Write-Host "$($Update.Title)" -ForegroundColor Cyan
                Write-Host "$($Update.OriginUri)" -ForegroundColor Gray
                Write-Host "$DownloadDirectory\$UpdateFile" -ForegroundColor Gray
                Start-BitsTransfer -Source $($Update.OriginUri) -Destination "$DownloadDirectory\$UpdateFile"
            }
        }
    }
    #===================================================================================================
    #   Export Install Script
    #===================================================================================================
    Write-Host "Update Install Script $PackagePath\Install-OSDUpdatePackage.ps1" -ForegroundColor Green
    Copy-Item "$($MyInvocation.MyCommand.Module.ModuleBase)\Scripts\Install-OSDUpdatePackage.ps1" "$PackagePath" -Force | Out-Null
    $ExportLine = "New-OSDUpdatePackage -PackageName '$PackageName' -PackagePath ""`$PSScriptRoot"" -RemoveSuperseded"
    $ExportLine | Out-File -FilePath "$PackagePath\Update-OSDUpdatePackage.ps1"
}