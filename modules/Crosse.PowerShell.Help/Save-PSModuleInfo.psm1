function Save-PSModuleInfo {
    [CmdletBinding()]
    param (
            [Parameter(Mandatory=$true)]
            [string]
            $DestinationPath
          )

    BEGIN {
        if ((Test-Path $DestinationPath) -eq $false) {
            mkdir $DestinationPath
        }
    }
    PROCESS {
        $ovp = $VerbosePreference
        $VerbosePreference = "SilentlyContinue"
        $modules = Get-Module -ListAvailable | Where-Object { $_.ModuleBase -match '.:\\(Windows\\system32\\WindowsPowerShell|Program Files)' }
        $c = $modules.Count
        for ($i = 0; $i -lt $c; $i++) {
            $m = $modules[$i]
            if ($m.Guid -eq [Guid]::Empty) {
                Write-Warning "GUID for $($m.Name) is empty! It will not be exported."
                continue
            }
            if ([String]::IsNullOrEmpty($m.HelpInfoUri)) {
                Write-Warning "No HelpInfoUri for $($m.Name). It will not be exported."
                continue
            }
            Write-Progress -Activity "Saving Modules to $DestinationPath" `
                           -Status " " `
                           -CurrentOperation $m.Name `
                           -PercentComplete ([Int32]([Math]::Floor($i/$c * 100)))
            $savePath = Join-Path $DestinationPath "$($m.Guid).xml"
            Export-Clixml -InputObject $m -Path $savePath
        }
    }
}
