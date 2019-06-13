<#
.SYNOPSIS
Downloads Microsoft Office 365 and Office 2019

.DESCRIPTION
Downloads Microsoft Office 365 and Office 2019
Requires Internet access for downloading the updates

.LINK
https://www.osdeploy.com/osdupdate/docs/functions/get-downoffice

.PARAMETER RepositoryRootPath
Full Path of the OSDUpdate Repository

.PARAMETER OfficeProduct
'Office 365 ProPlus','Office 365 Business','Office 2019 ProPlus','Office 2019 Standard'
Defaults to Office 365 ProPlus

.PARAMETER UpdateChannel
Insiders, Monthly, SAC, SACT
Defaults to SAC

.PARAMETER OfficeArch
32 or 64
Defaults to 64

.PARAMETER LanguageId
Defaults to en-us

.PARAMETER ProofingTools
Defaults to en-us

.PARAMETER XmlOnly
Creates an XML for use with ODT to manually download later
#>

function Get-DownOffice {
    [CmdletBinding()]
    PARAM (
        [Parameter(Mandatory = $True)]
        [string]$RepositoryRootPath,
        #===================================================================================================
        #   Office Download
        #===================================================================================================
        [ValidateSet(
            'Office 365 ProPlus',
            'Office 365 Business',
            'Office 2019 ProPlus',
            'Office 2019 Standard')]
        [string]$OfficeProduct = 'Office 365 ProPlus',

        [ValidateSet(
            'Insiders',
            'Monthly',
            'SAC',
            'SACT'
        )]
        [string]$UpdateChannel = 'SAC',

        [ValidateSet('64','32')]
        [string]$OfficeArch = '64',
        #===================================================================================================
        #   LanguageId
        #   https://docs.microsoft.com/en-us/DeployOffice/overview-of-deploying-languages-in-office-365-proplus
        #===================================================================================================
        [ValidateSet(
            'af-za',
            'ar-sa',
            'as-in',
            'az-Latn-az',
            'bg-bg',
            'bn-bd',
            'bn-in',
            'bs-latn-ba',
            'ca-es',
            'ca-es-valencia',
            'cs-cz',
            'Culture (ll-cc)',
            'cy-gb',
            'da-dk',
            'de-de',
            'el-gr',
            'en-us',
            'es-es',
            'et-ee',
            'eu-es',
            'fa-ir',
            'fi-fi',
            'fr-fr',
            'ga-ie',
            'gd-gb',
            'gl-es',
            'gu-in',
            'ha-Latn-ng',
            'he-il',
            'hi-in',
            'hr-hr',
            'hu-hu',
            'hy-am',
            'id-id',
            'ig-ng',
            'is-is',
            'it-it',
            'ja-jp',
            'ka-ge',
            'kk-kz',
            'kn-in',
            'kok-in',
            'ko-kr',
            'ky-kg',
            'lb-lu',
            'lt-lt',
            'lv-lv',
            'mi-nz',
            'mk-mk',
            'ml-in',
            'mr-in',
            'ms-my',
            'mt-mt',
            'nb-no',
            'ne-np',
            'nl-nl',
            'nn-no',
            'nso-za',
            'or-in',
            'pa-in',
            'pl-pl',
            'ps-af',
            'pt-br',
            'pt-pt',
            'rm-ch',
            'ro-ro',
            'ru-ru',
            'rw-rw',
            'si-lk',
            'sk-sk',
            'sl-si',
            'sq-al',
            'sr-cyrl-ba',
            'sr-cyrl-rs',
            'sr-latn-rs',
            'sv-se',
            'sw-ke',
            'ta-in',
            'te-in',
            'th-th',
            'tn-za',
            'tr-tr',
            'tt-ru',
            'uk-ua',
            'ur-pk',
            'uz-Latn-uz',
            'vi-vn',
            'wo-sn',
            'xh-za',
            'yo-ng',
            'zh-cn',
            'zh-tw',
            'zu-za')]
        [string]$LanguageId = 'en-US',
        [ValidateSet(
            'af-za',
            'ar-sa',
            'as-in',
            'az-Latn-az',
            'bg-bg',
            'bn-bd',
            'bn-in',
            'bs-latn-ba',
            'ca-es',
            'ca-es-valencia',
            'cs-cz',
            'Culture (ll-cc)',
            'cy-gb',
            'da-dk',
            'de-de',
            'el-gr',
            'en-us',
            'es-es',
            'et-ee',
            'eu-es',
            'fa-ir',
            'fi-fi',
            'fr-fr',
            'ga-ie',
            'gd-gb',
            'gl-es',
            'gu-in',
            'ha-Latn-ng',
            'he-il',
            'hi-in',
            'hr-hr',
            'hu-hu',
            'hy-am',
            'id-id',
            'ig-ng',
            'is-is',
            'it-it',
            'ja-jp',
            'ka-ge',
            'kk-kz',
            'kn-in',
            'kok-in',
            'ko-kr',
            'ky-kg',
            'lb-lu',
            'lt-lt',
            'lv-lv',
            'mi-nz',
            'mk-mk',
            'ml-in',
            'mr-in',
            'ms-my',
            'mt-mt',
            'nb-no',
            'ne-np',
            'nl-nl',
            'nn-no',
            'nso-za',
            'or-in',
            'pa-in',
            'pl-pl',
            'ps-af',
            'pt-br',
            'pt-pt',
            'rm-ch',
            'ro-ro',
            'ru-ru',
            'rw-rw',
            'si-lk',
            'sk-sk',
            'sl-si',
            'sq-al',
            'sr-cyrl-ba',
            'sr-cyrl-rs',
            'sr-latn-rs',
            'sv-se',
            'sw-ke',
            'ta-in',
            'te-in',
            'th-th',
            'tn-za',
            'tr-tr',
            'tt-ru',
            'uk-ua',
            'ur-pk',
            'uz-Latn-uz',
            'vi-vn',
            'wo-sn',
            'xh-za',
            'yo-ng',
            'zh-cn',
            'zh-tw',
            'zu-za')]
        [string]$ProofingTools = 'en-US',
        [switch]$XmlOnly
    )

    #===================================================================================================
    #   OfficeODT
    #===================================================================================================
    Write-Host '========================================================================================' -ForegroundColor DarkGray
    Write-Verbose 'Office Deployment Tool: https://www.microsoft.com/en-us/download/details.aspx?id=49117' -Verbose
    #$OfficeODTUrl = 'https://download.microsoft.com/download/2/7/A/27AF1BE6-DD20-4CB4-B154-EBAB8A7D4A7E/officedeploymenttool_11509-33604.exe'
    $OfficeODTUrl = 'https://download.microsoft.com/download/2/7/A/27AF1BE6-DD20-4CB4-B154-EBAB8A7D4A7E/officedeploymenttool_11617-33601.exe'
    $OfficeODTDir = "$RepositoryRootPath\OfficeODT"
    $OfficeODTDownloadFile = "officedeploymenttool_11617-33601.exe"
    $OfficeODT = 'setup.exe'
    $OfficeODTFullName = "$OfficeODTDir\$OfficeODT"
    if (!(Test-Path "$OfficeODTFullName")) {
        Write-Warning "Office ODT must be downloaded and extracted to $OfficeODTFullName"
        Write-Verbose "Office ODT Url: $OfficeODTUrl" -Verbose
        Write-Verbose "Office ODT Download File: $OfficeODTDownloadFile" -Verbose
        Write-Verbose "Office ODT Directory: $OfficeODTDir" -Verbose
        Write-Verbose "Office ODT: $OfficeODTFullName" -Verbose
        if (!(Test-Path "$OfficeODTDir")) {New-Item "$OfficeODTDir" -ItemType Directory -Force | Out-Null}
        Invoke-WebRequest -Uri $OfficeODTUrl -OutFile "$OfficeODTDir\$OfficeODTDownloadFile"
        Start-Process "$OfficeODTDir\$OfficeODTDownloadFile" -ArgumentList "/extract:`"$OfficeODTDir`"" -Wait
    }
    if (!(Test-Path "$OfficeODTFullName")) {
        Write-Warning "You will need to download and extract the Office Deployment Tool to $OfficeODTFullName before using OSDUpdate OfficeDownload"
        Break
    }

    if ($OfficeProduct -eq 'Office 365 ProPlus') {$ODTOfficeProduct = 'O365ProPlusRetail'}
    if ($OfficeProduct -eq 'Office 365 Business') {$ODTOfficeProduct = 'O365BusinessRetail'}
    if ($OfficeProduct -eq 'Office 2019 ProPlus') {
        $ODTOfficeProduct = 'Standard2019Volume'
        $ODTChannel = 'PerpetualVL2019'
    }
    if ($OfficeProduct -eq 'Office 2019 Standard') {
        $ODTOfficeProduct = 'Standard2019Volume'
        $ODTChannel = 'PerpetualVL2019'
    }
    if ($OfficeProduct -eq 'Project 2019 Pro') {
        $ODTOfficeProduct = 'ProjectPro2019Volume '
        $ODTChannel = 'PerpetualVL2019'
    }
    if ($OfficeProduct -eq 'Project 2019 Standard') {
        $ODTOfficeProduct = 'ProjectStd2019Volume'
        $ODTChannel = 'PerpetualVL2019'
    }
    if ($OfficeProduct -eq 'Visio 2019 Pro') {
        $ODTOfficeProduct = 'VisioPro2019Volume'
        $ODTChannel = 'PerpetualVL2019'
    }
    if ($OfficeProduct -eq 'Visio 2019 Standard') {
        $ODTOfficeProduct = 'VisioStd2019Volume'
        $ODTChannel = 'PerpetualVL2019'
    }

    if (($ODTOfficeProduct -eq 'O365ProPlusRetail') -or ($ODTOfficeProduct -eq 'O365BusinessRetail')) {
        if ($UpdateChannel -eq 'Insiders') {$ODTChannel = 'Insiders'}
        if ($UpdateChannel -eq 'Monthly') {$ODTChannel = 'Monthly'}
        if ($UpdateChannel -eq 'SAC') {$ODTChannel = 'Broad'}
        if ($UpdateChannel -eq 'SACT') {$ODTChannel = 'Targeted'}
    }
    #===================================================================================================
    #   OfficeDownload
    #===================================================================================================
    if ($UpdateChannel -eq 'Monthly') {
        $OfficeDownload = "$OfficeProduct $OfficeArch-Bit"
    } else {
        $OfficeDownload = "$OfficeProduct $OfficeArch-Bit $UpdateChannel"
    }
    $DownloadsPath = "$RepositoryRootPath\$OfficeDownload"

$ODTXml = @"
<Configuration>
<Add SourcePath="$DownloadsPath" OfficeClientEdition="$OfficeArch" Channel="$ODTChannel">
    <Product ID="$ODTOfficeProduct">
        <Language ID="$LanguageId" />
    </Product>
    <Product ID="ProofingTools">
        <Language ID="$ProofingTools" />
    </Product>
</Add>
</Configuration>
"@

    if (!(Test-Path $DownloadsPath)) {New-Item $DownloadsPath -ItemType Directory -Force | Out-Null}
    $ODTXml | Out-File "$DownloadsPath\OSDUpdate.xml" -Encoding utf8 -Force
    Write-Host '========================================================================================' -ForegroundColor DarkGray
    Write-Verbose "Office Deployment Tool Download XML saved to $DownloadsPath\OSDUpdate.xml" -Verbose
    Write-Verbose "You can download $OfficeProduct manually with the following command line:" -Verbose
    Write-Verbose "`"$OfficeODTFullName`" /download `"$DownloadsPath\OSDUpdate.xml`"" -Verbose
    Write-Host '========================================================================================' -ForegroundColor DarkGray
    Write-Verbose "ODT SourcePath: $DownloadsPath" -Verbose
    Write-Verbose "ODT OfficeClientEdition: $OfficeArch" -Verbose
    Write-Verbose "ODT Channel: $ODTChannel" -Verbose
    Write-Verbose "ODT Product Id: $ODTOfficeProduct" -Verbose
    Write-Verbose "ODT Language Primary: $LanguageId" -Verbose
    Write-Verbose "ODT Proofing Tools: $ProofingTools" -Verbose
    Write-Verbose "Download Full Path: $DownloadsPath" -Verbose
    Write-Host '========================================================================================' -ForegroundColor DarkGray
    if (!($XmlOnly)) {
        Write-Verbose "Downloading ... This may take a while ..." -Verbose
        Start-Process -FilePath "$OfficeODTFullName" -ArgumentList "/download","`"$DownloadsPath\OSDUpdate.xml`"" -Wait
        Write-Host '========================================================================================' -ForegroundColor DarkGray
    }
    Write-Verbose "Complete!" -Verbose
}