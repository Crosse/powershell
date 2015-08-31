@{
# Script module or binary module file associated with this manifest
ModuleToProcess = ''

# Version number of this module.
ModuleVersion = '1.0'

# ID used to uniquely identify this module
GUID = '696ffc46-85fa-4022-a538-db4ef4978927'

# Author of this module
Author = 'Seth Wright'

# Company or vendor of this module
CompanyName = 'James Madison University'

# Copyright statement for this module
Copyright = 'Copyright © 2011 Seth Wright <wrightst@jmu.edu>'

# Description of the functionality provided by this module
Description = 'Generic functions'

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
RequiredAssemblies = @('System.Speech')

# Script files (.psm1) that are run in the caller's environment prior to importing this module
ScriptsToProcess = @()

# Type files (.psm1xml) to be loaded when importing this module
TypesToProcess = @()

# Format files (.psm1xml) to be loaded when importing this module
FormatsToProcess = @()

# Modules to import as nested modules of the module specified in ModuleToProcess
NestedModules = 'DnsFunctions.psm1',
                'GeneralFunctions.psm1',
                'Publish-Item.psm1',
                'TextManipulation.psm1',
                'Base64Functions.psm1',
                'Get-DirectoryStatistics.psm1',
                'EventLogSummary.psm1',
                'Speech.psm1',
                'Connect-RemoteServer.psm1',
                'JobControl.psm1',
                'Invoke-TimedScriptBlock.psm1'

# Functions to export from this module
FunctionsToExport = '*'

# Cmdlets to export from this module
CmdletsToExport = '*'

# Variables to export from this module
VariablesToExport = '*'

# Aliases to export from this module
AliasesToExport = '*'

# List of all modules packaged with this module
ModuleList =    'DnsFunctions.psm1',
                'GeneralFunctions.psm1',
                'Publish-Item.psm1',
                'TextManipulation.psm1',
                'Base64Functions.psm1',
                'Get-DirectoryStatistics.psm1',
                'EventLogSummary.psm1',
                'Speech.psm1',
                'Connect-RemoteServer.psm1',
                'JobControl.psm1',
                'Invoke-TimedScriptBlock.psm1'

# List of all files packaged with this module
FileList =      'DnsFunctions.psm1',
                'GeneralFunctions.psm1',
                'Publish-Item.psm1',
                'TextManipulation.psm1',
                'Base64Functions.psm1',
                'Get-DirectoryStatistics.psm1',
                'EventLogSummary.psm1',
                'Speech.psm1',
                'Connect-RemoteServer.psm1',
                'JobControl.psm1',
                'Invoke-TimedScriptBlock.psm1'

# Private data to pass to the module specified in ModuleToProcess
PrivateData = ''

}
