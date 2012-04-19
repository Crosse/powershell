function Get-InstalledPatch {
    [CmdletBinding()]
    param (
            [Parameter(Mandatory=$false,
                ValueFromPipeline=$true)]
            [ValidateNotNullOrEmpty()]
            [string]
            $ComputerName = "localhost",

            [Parameter(Mandatory=$false)]
            [ValidateNotNullOrEmpty()]
            [string[]]
            $HotFixID
        )

    PROCESS {
        $os = Get-WmiObject -ComputerName $ComputerName `
                            -Class Win32_OperatingSystem `

        $hotfixes = @()
        if ([String]::IsNullOrEmpty($HotFixID)) {
            Write-Verbose "Getting all HotFixIDs"
            $hotfixes = Get-WmiObject -ComputerName $ComputerName -Class Win32_QuickFixEngineering
        } else {
            foreach ($hfid in $HotFixID) {
                $hotfix = Get-WmiObject -ComputerName $ComputerName `
                                -Class Win32_QuickFixEngineering `
                                -Filter "HotFixID = '$hfid'"

                if ($hotfix -eq $null) {
                    Write-Verbose "Did not find HotFixID $hfid"
                    New-Object PSObject -Property @{
                        "ComputerName"  = $ComputerName
                        "HotFixID"      = $hfid
                        "Installed"     = $false
                        "Description"   = $null
                        "InstalledOn"   = $null
                    }
                } else {
                    Write-Verbose "Found HotFixID $($hotfix.HotFixId)"
                    $hotfixes += $hotfix
                }
            }
        }

        foreach ($hotfix in $hotfixes) {
            New-Object PSObject -Property @{
                "ComputerName"  = $ComputerName
                "HotFixID"      = $hotfix.HotFixID
                "Installed"     = $true
                "Description"   = $hotfix.Description
                "InstalledOn"   = $hotfix.InstalledOn
            }
        }
    }
}
