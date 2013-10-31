#requires -version 3

workflow Get-InstalledPatch {
    [CmdletBinding()]
    param (
            [Parameter(Mandatory=$false)]
            [ValidateNotNullOrEmpty()]
            [string[]]
            $ComputerName,

            [Parameter(Mandatory=$true)]
            [ValidateNotNullOrEmpty()]
            [string]
            $HotFixID
        )

    foreach -parallel ($Computer in $ComputerName) {
        Write-Verbose -Message "Searching for HotFix $HotFixID on $Computer"
        $hotfix = Get-WmiObject -PSComputerName $Computer `
                    -Class Win32_QuickFixEngineering `
                    -Filter "HotFixID = '$HotFixID'"

        if ($hotfix -eq $null -or [String]::IsNullOrEmpty($hotfix)) {
            Write-Verbose -Message "Did not find HotFixID $HotFixID on $Computer"
        } else {
            Write-Verbose -Message "Found HotFixID $($hotfix.HotFixId) on $Computer"
            $p2 = New-Object -TypeName PSObject -Property @{
                "ComputerName"  = $Computer
                "HotFixID"      = $hotfix.HotFixID
                "Installed"     = $true
                "Description"   = $hotfix.Description
                "InstalledOn"   = $hotfix.InstalledOn
            }
            Write-Output -InputObject $p2
        }
    }
}
