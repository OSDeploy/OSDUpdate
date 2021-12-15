# Module Manifest
#

@{

# Script module or binary module file associated with this manifest.
RootModule = 'OSDUpdate.psm1'

# Version number of his module.
ModuleVersion = '21.12.14.1'

# Supported PSEditions
# CompatiblePSEditions = @()

# ID used to uniquely identify this module
GUID = '593e8730-b646-4a16-abaf-f3eae38c57b0'

# Author of this module
Author = 'David Segura'

# Company or vendor of this module
CompanyName = 'osdeploy.com'

# Copyright statement for this module
Copyright = '(c) 2021 David Segura osdeploy.com. All rights reserved.'

# Description of the functionality provided by this module
Description = @'
Requires OSDSUS 21.10.14.1 or newer

OSDUpdate https://osdupdate.osdeploy.com/

Latest Microsoft Updates:
https://raw.githubusercontent.com/OSDeploy/OSDUpdate/master/UPDATES.md

WSUS Update Catalogs:
These are contained within this PowerShell Module, so regular Module updating is needed to
ensure you receive the latest Microsoft Updates.  Updates published in WSUS will be different
from Microsoft Update Catalog website due to Preview Releases
'@

# Minimum version of the Windows PowerShell engine required by this module
PowerShellVersion = '5.0'

# Name of the Windows PowerShell host required by this module
# PowerShellHostName = 'Windows PowerShell ISE Host'

# Minimum version of the Windows PowerShell host required by this module
# PowerShellHostVersion = ''

# Minimum version of Microsoft .NET Framework required by this module. This prerequisite is valid for the PowerShell Desktop edition only.
# DotNetFrameworkVersion = ''

# Minimum version of the common language runtime (CLR) required by this module. This prerequisite is valid for the PowerShell Desktop edition only.
# CLRVersion = ''

# Processor architecture (None, X86, Amd64) required by this module
# ProcessorArchitecture = ''

# Modules that must be imported into the global environment prior to importing this module
RequiredModules = @(
    @{ModuleName='OSDSUS'; ModuleVersion = '21.10.14.1'; Guid="065cf035-da73-4d17-8745-f55116b82fb5"}
)

# Assemblies that must be loaded prior to importing this module
# RequiredAssemblies = @()

# Script files (.ps1) that are run in the caller's environment prior to importing this module.
# ScriptsToProcess = @()

# Type files (.ps1xml) to be loaded when importing this module
# TypesToProcess = @()

# Format files (.ps1xml) to be loaded when importing this module
# FormatsToProcess = @()

# Modules to import as nested modules of the module specified in RootModule/ModuleToProcess
#NestedModules = @(
#    @{ModuleName="OSDSUS"; ModuleVersion = '19.9.10.0'; Guid="065cf035-da73-4d17-8745-f55116b82fb5"}
#)

# Functions to export from this module, for best performance, do not use wildcards and do not delete the entry, use an empty array if there are no functions to export.
FunctionsToExport = 'Get-OSDUpdate',
                    'Get-DownDefender',
                    'Get-DownMcAfee',
                    'Get-DownOSDUpdate',
                    'Get-DownOffice',
                    'New-OSDUpdatePackage',
                    #'New-OSDUpdateRepository',
                    'Update-ModuleOSDUpdate'

# Cmdlets to export from this module, for best performance, do not use wildcards and do not delete the entry, use an empty array if there are no cmdlets to export.
CmdletsToExport = @()

# Variables to export from this module
VariablesToExport = @()

# Aliases to export from this module, for best performance, do not use wildcards and do not delete the entry, use an empty array if there are no aliases to export.
AliasesToExport = @()

# DSC resources to export from this module
# DscResourcesToExport = @()

# List of all modules packaged with this module
# ModuleList = @()

# List of all files packaged with this module
# FileList = @()

# Private data to pass to the module specified in RootModule/ModuleToProcess. This may also contain a PSData hashtable with additional module metadata used by PowerShell.
PrivateData = @{

    PSData = @{

        # Tags applied to this module. These help with module discovery in online galleries.
        Tags = @('OSDeploy','OSDUpdate','OSDSUS','OSD','Update','Windows10','Office365','Office2019','Office2016','Office2013','Office2010')

        # A URL to the license for this module.
        LicenseUri = 'https://github.com/OSDeploy/OSDUpdate/blob/master/LICENSE'

        # A URL to the main website for this project.
        ProjectUri = 'https://osdupdate.osdeploy.com/'

        # A URL to an icon representing this module.
        IconUri = 'https://raw.githubusercontent.com/OSDeploy/OSDUpdate/master/OSD.png'

        # ReleaseNotes of this module
        ReleaseNotes = 'https://osdupdate.osdeploy.com/release'

        #ExternalModuleDependencies = @('OSDSUS')

    } # End of PSData hashtable

} # End of PrivateData hashtable

# HelpInfo URI of this module
# HelpInfoURI = ''

# Default prefix for commands exported from this module. Override the default prefix using Import-Module -Prefix.
# DefaultCommandPrefix = ''

}
