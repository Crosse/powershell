﻿@{

# Script module or binary module file associated with this manifest
ModuleToProcess = ''

# Version number of this module.
ModuleVersion = '1.0'

# ID used to uniquely identify this module
GUID = '19212ba6-53f4-430a-bc7f-ed56de9df60d'

# Author of this module
Author = 'Seth Wright'

# Company or vendor of this module
CompanyName = 'James Madison University'

# Copyright statement for this module
Copyright = 'Copyright © 2015 Seth Wright <wrightst@jmu.edu>'

# Description of the functionality provided by this module
Description = 'Cmdlets that may only be useful to JMU'

# Minimum version of the Windows PowerShell engine required by this module
PowerShellVersion = '2.0'

# Name of the Windows PowerShell host required by this module
PowerShellHostName = ''

# Minimum version of the Windows PowerShell host required by this module
PowerShellHostVersion = ''

# Minimum version of the .NET Framework required by this module
DotNetFrameworkVersion = ''

# Minimum version of the common language runtime (CLR) required by this module
CLRVersion = ''

# Processor architecture (None, X86, Amd64, IA64) required by this module
ProcessorArchitecture = ''

# Modules that must be imported into the global environment prior to importing this module
RequiredModules = @()

# Assemblies that must be loaded prior to importing this module
RequiredAssemblies = @()

# Script files (.ps1) that are run in the caller's environment prior to importing this module
ScriptsToProcess = @()

# Type files (.ps1xml) to be loaded when importing this module
TypesToProcess = @()

# Format files (.ps1xml) to be loaded when importing this module
FormatsToProcess = @()

# Modules to import as nested modules of the module specified in ModuleToProcess
NestedModules = 'Search-LockoutEvents.psm1',
                'Get-LockedUser.psm1',
                'Search-ACSFailedAuthLogs.psm1',
                'Get-AuthenticationFailures.psm1',
                'Get-ReservedIdentifier.psm1',
                'New-RemedyTicket.psm1'

# Functions to export from this module
FunctionsToExport = '*'

# Cmdlets to export from this module
CmdletsToExport = '*'

# Variables to export from this module
VariablesToExport = '*'

# Aliases to export from this module
AliasesToExport = '*'

# List of all modules packaged with this module
ModuleList =    'Search-LockoutEvents.psm1',
                'Get-LockedUser.psm1',
                'Search-ACSFailedAuthLogs.psm1',
                'Get-AuthenticationFailures.psm1',
                'Get-ReservedIdentifier.psm1',
                'New-RemedyTicket.psm1'

# List of all files packaged with this module
FileList =      'Search-LockoutEvents.psm1',
                'Get-LockedUser.psm1',
                'Search-ACSFailedAuthLogs.psm1',
                'Get-AuthenticationFailures.psm1',
                'Get-ReservedIdentifier.psm1',
                'New-RemedyTicket.psm1'

# Private data to pass to the module specified in ModuleToProcess
PrivateData = ''

}

