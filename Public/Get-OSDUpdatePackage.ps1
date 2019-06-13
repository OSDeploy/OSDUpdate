function Get-OSDUpdatePackage {
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
        
        [Parameter(Mandatory = $True)]
        [string]$PackagePath,
        [switch]$AppendPackageName,
        [switch]$RemoveSuperseded,

        [ValidateSet('SeguraOSD','SuneThomsenDK')]
        [string]$InstallScript = 'SeguraOSD',

        [ValidateSet('Default','Proofing','Language','All')]
        [string]$OfficeProfile = 'Default',

        [Alias('OfficeSetupUpdatesPath')]
        [string]$OfficeMediaUpdatesPath,

        [switch]$GridView
    )
    #===================================================================================================
    #   Variables
    #===================================================================================================
    $AllOSDUpdates = @()

    if ($AppendPackageName) {
        $PackagePath = "$PackagePath\$PackageName"
    }
    #===================================================================================================
    #   PackagePath
    #===================================================================================================
    if (!(Test-Path "$PackagePath")) {New-Item -Path "$PackagePath" -ItemType Directory -Force | Out-Null}
    #===================================================================================================
    #   Catalogs
    #===================================================================================================
    if ($PackageName -like "Office*") {
        $AllOnlineUpdates = Get-OSDUpdate -Catalog $PackageName -OfficeProfile All -Silent
        $AllOSDUpdates = Get-OSDUpdate -Catalog $PackageName -OfficeProfile $OfficeProfile

        Copy-Item -Path "$($MyInvocation.MyCommand.Module.ModuleBase)\Catalogs\$PackageName.xml" -Destination "$PackagePath" -Force | Out-Null
    }

    if ($PackageName -like "*Windows*") {
        if ($PackageName -like "Windows 7*") {
            $AllOnlineUpdates = Get-OSDUpdate -Catalog 'Windows 7'
            Copy-Item -Path "$($MyInvocation.MyCommand.Module.ModuleBase)\Catalogs\Windows 7.xml" -Destination "$PackagePath" -Force | Out-Null

        } elseif ($PackageName -like "Windows 10*") {
            $AllOnlineUpdates = Get-OSDUpdate -Catalog 'Windows 10'
            Copy-Item -Path "$($MyInvocation.MyCommand.Module.ModuleBase)\Catalogs\Windows 10.xml" -Destination "$PackagePath" -Force | Out-Null

        } elseif ($PackageName -like "Windows Server 2016*") {
            $AllOnlineUpdates = Get-OSDUpdate -Catalog 'Windows Server 2016'
            Copy-Item -Path "$($MyInvocation.MyCommand.Module.ModuleBase)\Catalogs\Windows Server 2016.xml" -Destination "$PackagePath" -Force | Out-Null

        } elseif ($PackageName -like "Windows Server 2019*") {
            $AllOnlineUpdates = Get-OSDUpdate -Catalog 'Windows Server 2019'
            Copy-Item -Path "$($MyInvocation.MyCommand.Module.ModuleBase)\Catalogs\Windows Server 2019.xml" -Destination "$PackagePath" -Force | Out-Null
        }

        $AllOSDUpdates = $AllOnlineUpdates
    }
    #===================================================================================================
    #   Existing Updates
    #===================================================================================================
    $ExistingUpdates = @()
    $SupersededUpdates = @()

    $ExistingUpdates = Get-ChildItem -Path "$PackagePath\*" -Directory -Recurse | Select-Object -Property *

    foreach ($Update in $ExistingUpdates) {
        if ($AllOnlineUpdates.Title -NotContains $Update.Name) {$SupersededUpdates += $Update.FullName}
    }
    #===================================================================================================
    #   Superseded Updates
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
    #   Get Downloaded Updates
    #===================================================================================================
    foreach ($Update in $AllOSDUpdates) {
        if ($PackageName -like "Windows*") {
            $FullUpdatePath = "$PackagePath\$($Update.Title)\$($Update.FileName)"
            if (Test-Path $FullUpdatePath) {
                $Update.OSDStatus = "Downloaded"
            }
        }
    }
    #===================================================================================================
    #   Filters
    #===================================================================================================
    if ($PackageName -like "Windows*") {
        if ($PackageName -like "*x64*") {$AllOSDUpdates = $AllOSDUpdates | Where-Object {$_.UpdateArch -eq 'x64'}}
        if ($PackageName -like "*x86*") {$AllOSDUpdates = $AllOSDUpdates | Where-Object {$_.UpdateArch -eq 'x86'}}

        if ($PackageName -like "*1507*") {$AllOSDUpdates = $AllOSDUpdates | Where-Object {$_.UpdateBuild -eq '1507'}}
        if ($PackageName -like "*1511*") {$AllOSDUpdates = $AllOSDUpdates | Where-Object {$_.UpdateBuild -eq '1511'}}
        if ($PackageName -like "*1607*") {$AllOSDUpdates = $AllOSDUpdates | Where-Object {$_.UpdateBuild -eq '1607'}}
        if ($PackageName -like "*1703*") {$AllOSDUpdates = $AllOSDUpdates | Where-Object {$_.UpdateBuild -eq '1703'}}
        if ($PackageName -like "*1709*") {$AllOSDUpdates = $AllOSDUpdates | Where-Object {$_.UpdateBuild -eq '1709'}}
        if ($PackageName -like "*1803*") {$AllOSDUpdates = $AllOSDUpdates | Where-Object {$_.UpdateBuild -eq '1803'}}
        if ($PackageName -like "*1809*") {$AllOSDUpdates = $AllOSDUpdates | Where-Object {$_.UpdateBuild -eq '1809'}}
        if ($PackageName -like "*1903*") {$AllOSDUpdates = $AllOSDUpdates | Where-Object {$_.UpdateBuild -eq '1903'}}
    }
    #===================================================================================================
    #   GridView
    #===================================================================================================
    if ($PackageName -like "Office*") {
        $AllOSDUpdates = $AllOSDUpdates | Select-Object -Property OSDStatus,Catalog,CreationDate,KBNumber,Title,FileName,Size,FileUri,OriginUri,OSDGuid
    }
    if ($PackageName -like "Windows*") {
        $AllOSDUpdates = $AllOSDUpdates | Select-Object -Property OSDStatus,Catalog,UpdateOS,UpdateArch,UpdateBuild,CreationDate,KBNumber,Title,FileName,Size,FileUri,OriginUri,OSDGuid
    }
    if ($GridView.IsPresent) {$AllOSDUpdates = $AllOSDUpdates | Out-GridView -PassThru -Title "Select OSDUpdate Downloads to include in the Package"}
    #===================================================================================================
    #   Sort
    #===================================================================================================
    $AllOSDUpdates = $AllOSDUpdates | Sort-Object DateCreated
    #===================================================================================================
    #   Download
    #===================================================================================================
    if ($PackageName -like "Office*") {
        foreach ($Update in $AllOSDUpdates) {
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

        Write-Verbose "Adding $PackagePath\OSDUpdate-Office.ps1" -Verbose
        Copy-Item "$($MyInvocation.MyCommand.Module.ModuleBase)\Scripts\OSDUpdate-Office.ps1" "$PackagePath" -Force | Out-Null
    }

    if ($PackageName -like "Windows*") {
        foreach ($Update in $AllOSDUpdates) {
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
        Write-Verbose "Adding $PackagePath\OSDUpdate-Windows.ps1" -Verbose
        Copy-Item "$($MyInvocation.MyCommand.Module.ModuleBase)\Scripts\OSDUpdate-Windows.ps1" "$PackagePath" -Force | Out-Null
    }
    #===================================================================================================
}