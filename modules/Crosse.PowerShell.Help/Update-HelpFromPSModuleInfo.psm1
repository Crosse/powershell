function Update-HelpFromPSModuleInfo {
    [CmdletBinding()]
    param (
            [Parameter(Mandatory=$true)]
            [string]
            $SourcePath,

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
        $modules = Get-ChildItem $SourcePath -Recurse -Filter *.xml | Import-Clixml
        $c = $modules.Count
        for ($i = 0; $i -lt $c; $i++) {
            $m = $modules[$i]
            if ($m.Guid -eq [Guid]::Empty) {
                Write-Warning "GUID for $($m.Name) is empty! It will not be imported."
                continue
            }
            if ([String]::IsNullOrEmpty($m.HelpInfoUri)) {
                Write-Warning "No HelpInfoUri for $($m.Name). It will not be imported."
                continue
            }
            Write-Progress -Activity "Updating Module Help" `
                           -Status " " `
                           -CurrentOperation $m.Name `
                           -PercentComplete ([Int32]([Math]::Floor($i/$c * 100)))
            Save-Help -Module $m -DestinationPath $DestinationPath
        }
    }
}
