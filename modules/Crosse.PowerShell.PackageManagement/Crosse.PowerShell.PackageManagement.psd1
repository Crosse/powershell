#
# Module manifest for module 'Crosse.PowerShell.PackageManagement'
#
# Generated by: Seth Wright
#
# Generated on: 3/1/2012
#
################################################################################
#
# Copyright (c) 2012 Seth Wright <wrightst@jmu.edu>
#
# Permission to use, copy, modify, and distribute this software for any
# purpose with or without fee is hereby granted, provided that the above
# copyright notice and this permission notice appear in all copies.
#
# THE SOFTWARE IS PROVIDED "AS IS" AND THE AUTHOR DISCLAIMS ALL WARRANTIES
# WITH REGARD TO THIS SOFTWARE INCLUDING ALL IMPLIED WARRANTIES OF
# MERCHANTABILITY AND FITNESS. IN NO EVENT SHALL THE AUTHOR BE LIABLE FOR
# ANY SPECIAL, DIRECT, INDIRECT, OR CONSEQUENTIAL DAMAGES OR ANY DAMAGES
# WHATSOEVER RESULTING FROM LOSS OF USE, DATA OR PROFITS, WHETHER IN AN
# ACTION OF CONTRACT, NEGLIGENCE OR OTHER TORTIOUS ACTION, ARISING OUT OF
# OR IN CONNECTION WITH THE USE OR PERFORMANCE OF THIS SOFTWARE.
#
################################################################################

@{

# Script module or binary module file associated with this manifest
# RootModule = ''

# Version number of this module.
ModuleVersion = '1.0'

# ID used to uniquely identify this module
GUID = '4ea9aa6b-6427-45d2-8ccc-394547b0eabb'

# Author of this module
Author = 'Seth Wright'

# Company or vendor of this module
CompanyName = 'James Madison University'

# Copyright statement for this module
Copyright = 'Copyright (c) 2012 Seth Wright <wrighst@jmu.edu>'

# Description of the functionality provided by this module
# Description = ''

# Minimum version of the Windows PowerShell engine required by this module
PowerShellVersion = '2.0'

# Name of the Windows PowerShell host required by this module
# PowerShellHostName = ''

# Minimum version of the Windows PowerShell host required by this module
# PowerShellHostVersion = ''

# Minimum version of the .NET Framework required by this module
DotNetFrameworkVersion = ''

# Minimum version of the common language runtime (CLR) required by this module
# CLRVersion = ''

# Processor architecture (None, X86, Amd64) required by this module
# ProcessorArchitecture = ''

# Modules that must be imported into the global environment prior to importing this module
# RequiredModules = @()

# Assemblies that must be loaded prior to importing this module
# RequiredAssemblies = @()

# Script files (.ps1) that are run in the caller's environment prior to importing this module
ScriptsToProcess = @('Startup.ps1')

# Type files (.ps1xml) to be loaded when importing this module
# TypesToProcess = @()

# Format files (.ps1xml) to be loaded when importing this module
FormatsToProcess = @('PackageManagement.Format.ps1xml')

# Modules to import as nested modules of the module specified in RootModule/ModuleToProcess
NestedModules = @(
    'Add-PackageItem.psm1',
    'Export-PackageItem.psm1',
    'Get-Package.psm1',
    'Get-PackageItem.psm1',
    'New-Package.psm1',
    'Out-Package.psm1',
    'Remove-PackageItem.psm1',
    'Set-Package.psm1'
                 )

# Functions to export from this module
FunctionsToExport = '*'

# Cmdlets to export from this module
CmdletsToExport = '*'

# Variables to export from this module
VariablesToExport = '*'

# Aliases to export from this module
AliasesToExport = '*'

# Commands to export from this module as Workflows
# ExportAsWorkflow = @()

# List of all modules packaged with this module
ModuleList = @(
    'Add-PackageItem.psm1',
    'Export-PackageItem.psm1',
    'Get-Package.psm1',
    'Get-PackageItem.psm1',
    'New-Package.psm1',
    'Out-Package.psm1',
    'Remove-PackageItem.psm1',
    'Set-Package.psm1'
              )

# List of all files packaged with this module
FileList = @(
    'Add-PackageItem.psm1',
    'Export-PackageItem.psm1',
    'Get-Package.psm1',
    'Get-PackageItem.psm1',
    'New-Package.psm1',
    'Out-Package.psm1',
    'Remove-PackageItem.psm1',
    'Set-Package.psm1'
            )

# Private data to pass to the module specified in RootModule/ModuleToProcess
# PrivateData = ''

# HelpInfo URI of this module
# HelpInfoURI = ''

# Default prefix for commands exported from this module. Override the default prefix using Import-Module -Prefix.
# DefaultCommandPrefix = ''

}

