function Get-AuthenticationFailures {
    [CmdletBinding()]
    param (
            [Parameter(Mandatory=$false)]
            [String[]]
            # The computer where the logs reside.  Defaults to the local machine.
            $ComputerName = $(Get-Content Env:\COMPUTERNAME).Value,

            [Parameter(Mandatory=$false)]
            [ValidateNotNullOrEmpty()]
            $EventLogName = "Security",

            [Parameter(Mandatory=$false)]
            [DateTime]
            # The starting time for the search.  The default is five minutes ago.
            $Start = (Get-Date).AddMinutes(-5),

            [Parameter(Mandatory=$false)]
            [DateTime]
            # The ending time for the search.  The default is the current time.
            $End = (Get-Date),

            [switch]
            $IncludeEvents = $false
          )

    BEGIN {
        $sw = New-Object System.Diagnostics.Stopwatch
        $startTime = $Start.ToUniversalTime().ToString('yyyy-MM-ddTHH:mm:ss.fffZ')
        $endTime = $End.ToUniversalTime().ToString('yyyy-MM-ddTHH:mm:ss.fffZ')

        $xPathQuery = @"
<QueryList>
<Query Id="0" Path="$EventLogName">
    <Select Path="$EventLogName">
        *[System[EventID=4625 and TimeCreated[@SystemTime &gt;= "$startTime" and @SystemTime &lt;= "$endTime"]]]
$userQuery
    </Select>
</Query>
</QueryList>
"@
    }
    PROCESS {
        $hits = @{}
        $total = $ComputerName.Count
        foreach ($c in $ComputerName) {
            $pctComplete = ($ComputerName.Count - $total) / $ComputerName.Count * 100
            $total--
            Write-Progress -Activity "Searching event logs" -Status "Querying $c" -PercentComplete $pctComplete
            Write-Verbose "Starting search on $c"
            $sw.Restart()
            $session = New-Object System.Diagnostics.Eventing.Reader.EventLogSession $c
            $logQuery = New-Object System.Diagnostics.Eventing.Reader.EventLogQuery $EventLogName, "LogName", $xPathQuery
            $logQuery.Session = $session
            $logReader = New-Object System.Diagnostics.Eventing.Reader.EventLogReader $logQuery

            $i = 0
            while (($evt = $logReader.ReadEvent()) -ne $null) {
                $props = Get-EventData -Event $evt
                $user = $props["TargetUserName"]
                if ($props) {
                    if ([String]::IsNullOrEmpty($user)) {
                        $user = "(none)"
                    }
                    if ($props["IpAddress"] -eq "-") {
                        $props["IpAddress"] = $null
                    }

                    if ($hits[$user] -eq $null) {
                        $hits[$user] = @()
                    }
                    $hits[$user] += $props
                }
            }
            $sw.Stop()
            Write-Verbose "Search ended on $c (took $($sw.ElapsedMilliseconds)ms)"
        }
        Write-Progress -Activity "Searching event logs" -Status "Done." -Completed

        Write-Verbose "Calculating Statistics"
        foreach ($hit in $hits.Keys) {
            $obj = New-Object PSObject -Property @{
                UserName                = $hit
                TotalFailedAuths        = $hits[$hit].Count
                FailedAuthsPerSecond    = [Math]::Round($hits[$hit].Count / ($End - $Start).TotalSeconds, 2)
                IPAddresses             = @($hits[$hit] | % { $_["IpAddress"] } | Sort -Unique)
            }
            if ($IncludeEvents) {
                $obj = Add-Member -InputObject $obj -MemberType NoteProperty `
                       -Name "Events" -Value @($hits[$hit]) -PassThru
            }
            $obj
        }
    }
}
