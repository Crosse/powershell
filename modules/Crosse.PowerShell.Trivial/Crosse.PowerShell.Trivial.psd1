@{

# Script module or binary module file associated with this manifest
ModuleToProcess = ''

# Version number of this module.
ModuleVersion = '1.0'

# ID used to uniquely identify this module
GUID = '644e2ae3-6006-4c6b-9a8e-8735f7363f5d'

# Author of this module
Author = 'Seth Wright'

# Company or vendor of this module
CompanyName = 'James Madison University'

# Copyright statement for this module
Copyright = 'Copyright © 2011 Seth Wright <wrightst@jmu.edu>'

# Description of the functionality provided by this module
Description = 'Trivial functions that do not belong anywhere else.'

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

# Script files (.psm1) that are run in the caller's environment prior to importing this module
ScriptsToProcess = @()

# Type files (.psm1xml) to be loaded when importing this module
TypesToProcess = @()

# Format files (.psm1xml) to be loaded when importing this module
FormatsToProcess = @()

# Modules to import as nested modules of the module specified in ModuleToProcess
NestedModules = 'GeoLocation.psm1',
                'Weather.psm1',
                'Get-ConsoleColors.psm1'

# Functions to export from this module
FunctionsToExport = '*'

# Cmdlets to export from this module
CmdletsToExport = '*'

# Variables to export from this module
VariablesToExport = '*'

# Aliases to export from this module
AliasesToExport = '*'

# List of all modules packaged with this module
ModuleList =    'GeoLocation.psm1',
                'Weather.psm1',
                'Get-ConsoleColors.psm1'

# List of all files packaged with this module
FileList =      'GeoLocation.psm1',
                'Weather.psm1',
                'Get-ConsoleColors.psm1'

# Private data to pass to the module specified in ModuleToProcess
PrivateData = ''
}
