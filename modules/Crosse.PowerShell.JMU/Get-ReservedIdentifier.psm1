function Get-ReservedIdentifier {
    [CmdletBinding()]
    param (
            [DateTime]
            $Start = ([DateTime]"1601-01-01 00:00:00.00"),

            [DateTime]
            $End = ([DateTime]::Now),

            [string[]]
            $Properties = @("cn", "sAMAccountName", "proxyAddresses", "mail"),

            [string]
            $ValidationRegex = '^((sip|smtp):)?(?<name>[a-zA-Z0-9]{4,8})(@.*)?$',

            [switch]
            $ShowDefaultValidationRegex
          )

    if ($ShowDefaultValidationRegex) {
        Write-Output $ValidationRegex
        return
    }

    $startfmt = $Start.ToString('yyyyMMddHHmmss.hhZ')
    $endfmt = $End.ToString('yyyyMMddHHmmss.hhZ')

    Write-Verbose "Getting objects created between $Start and $End"

    $obj = @(Get-ADObject -Filter { ((ObjectCategory -eq "person") -or (ObjectCategory -eq "group")) -and (WhenCreated -ge $startfmt) -and (WhenCreated -lt $endfmt) } -Properties $Properties)
    Write-Verbose "Found $($obj.Count) objects ($([DateTime]::Now))"

    $total = $obj.Count

    $reserved = @{}
    $sw = New-Object System.Diagnostics.Stopwatch
    $sw.Start()
    for ($i = 0; $i -lt $total; $i++) {
        $pctComplete = ($i/$total * 100)
        $o = $obj[$i]
        if ($i -ge 100 -and $i % 100) {
            $ticksPerUser = $sw.Elapsed.Ticks / $i
            $secondsRemaining = [TimeSpan]::FromTicks((($total - $i) * $ticksPerUser)).TotalSeconds
            Write-Progress -Activity "Scanning objects" -Status $o.Name -SecondsRemaining $secondsRemaining -PercentComplete ([Int32]$pctComplete)
        }
        foreach ($prop in $Properties) {
            foreach ($val in $o.$prop) {
                if ($val -match $validationRegex) {
                    $basename = $Matches["name"]
                    if ($reserved[$basename] -eq $null) {
                        $reserved[$basename] = ""
                        $basename
                    }
                }
            }
        }
    }
    $sw.Stop()
    Write-Progress -Activity "Scanning objects" -Status "Done" -Completed
    Write-Verbose "Done. ($([DateTime]::Now))"
}
