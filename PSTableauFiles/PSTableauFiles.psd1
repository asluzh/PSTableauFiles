#
# Module manifest for module 'PSTableauFiles'
#
# Generated by: Andrey Sluzhivoy
#
# Generated on: 8/30/2023
#

@{

# Script module or binary module file associated with this manifest.
RootModule = 'PSTableauFiles.psm1'

# Version number of this module.
ModuleVersion = '0.3.1'

# Supported PSEditions
# CompatiblePSEditions = @()

# ID used to uniquely identify this module
GUID = '868d1e0f-b6e8-4d56-9c3f-093744f13404'

# Author of this module
Author = 'Andrey Sluzhivoy'

# Company or vendor of this module
# CompanyName = 'D ONE'

# Copyright statement for this module
# Copyright = '(c) Andrey Sluzhivoy. All rights reserved.'

# Description of the functionality provided by this module
Description = 'This PowerShell module facilitates manipulating Tableau files for automation tasks.'
# Inspired by Tableau Document API
# https://github.com/tableau/document-api-python/tree/master/tableaudocumentapi

# Minimum version of the PowerShell engine required by this module
PowerShellVersion = '5.1'

# Modules that must be imported into the global environment prior to importing this module
# RequiredModules = @()

# Assemblies that must be loaded prior to importing this module
# RequiredAssemblies = @()

# Script files (.ps1) that are run in the caller's environment prior to importing this module.
# ScriptsToProcess = @()

# Type files (.ps1xml) to be loaded when importing this module
# TypesToProcess = @()

# Format files (.ps1xml) to be loaded when importing this module
# FormatsToProcess = @()

# Modules to import as nested modules of the module specified in RootModule/ModuleToProcess
# NestedModules = @()

# Functions to export from this module, for best performance, do not use wildcards and do not delete the entry, use an empty array if there are no functions to export.
FunctionsToExport = @(
'New-TableauZipFile', 'Update-TableauZipFile', 'Test-TableauZipFile',
'Get-TableauFileXml',
'Get-TableauFileStructure',
'Get-TableauFileObject' #, 'Export-TableauFileObject', 'Edit-TableauFileObject'
)

# Cmdlets to export from this module, for best performance, do not use wildcards and do not delete the entry, use an empty array if there are no cmdlets to export.
# CmdletsToExport = @()

# Variables to export from this module
# VariablesToExport = '*'

# Aliases to export from this module, for best performance, do not use wildcards and do not delete the entry, use an empty array if there are no aliases to export.
# AliasesToExport = @()

# DSC resources to export from this module
# DscResourcesToExport = @()

# List of all modules packaged with this module
# ModuleList = @()

# List of all files packaged with this module
FileList = @('PSTableauFiles.psm1', 'PSTableauFiles.psd1')

# Private data to pass to the module specified in RootModule/ModuleToProcess. This may also contain a PSData hashtable with additional module metadata used by PowerShell.
PrivateData = @{
    PSData = @{
        # Tags applied to this module. These help with module discovery in online galleries.
        Tags = @('tableau','xml','tableauworkbook')

        # A URL to the license for this module.
        LicenseUri = 'https://github.com/asluzh/PSTableauFiles/blob/main/LICENSE'

        # A URL to the main website for this project.
        ProjectUri = 'https://github.com/asluzh/PSTableauFiles'

        # A URL to an icon representing this module.
        # IconUri = ''
    }
}

# HelpInfo URI of this module
# HelpInfoURI = ''

# Default prefix for commands exported from this module. Override the default prefix using Import-Module -Prefix.
# DefaultCommandPrefix = ''

}
