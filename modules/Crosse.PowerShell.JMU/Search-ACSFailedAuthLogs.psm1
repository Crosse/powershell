function Search-ACSFailedAuthLogs {
    [CmdletBinding()]
    param (
            [Parameter(Mandatory=$true, ValueFromPipeline=$true)]
            [String]
            # The user to search for.  If no user is specified, all events will be returned.
            $User,

            [Parameter(Mandatory=$false)]
            [DateTime]
            # The starting time for the search.  The default is one hour ago.
            $Start = (Get-Date).AddHours(-1),

            [Parameter(Mandatory=$false)]
            [DateTime]
            # The ending time for the search.  The default is the current time.
            $End = (Get-Date),

            [Parameter(Mandatory=$false)]
            [String[]]
            # The computer where the logs reside.  Defaults to the local machine.
            $ComputerName = $(Get-Content Env:\COMPUTERNAME),

            [Parameter(Mandatory=$false)]
            # The SMB path to the directory containing ACS "Failed Attempts" log files.
            [String]
            $LogPath,

            [Parameter(Mandatory=$false)]
            # The number of events to return.  By default, all events are returned.
            [Long]
            $ResultSize = [Long]::MaxValue,

            [Parameter(Mandatory=$false)]
            # The path on the local computer to the Microsoft LogParser utility.
            [String]
            $LogParserPath = "$(Get-Content 'Env:\ProgramFiles(x86)')\Log Parser 2.2\LogParser.exe"
          )

    BEGIN {
        $query = @"
SELECT
     [Caller-ID] AS MACAddress
    ,TO_TIMESTAMP(TO_TIMESTAMP(Date, 'MM/dd/yyyy'), TO_TIMESTAMP(Time, 'HH:mm:ss')) AS Timestamp
    ,[User-Name] AS UserName
    ,[Group-Name] AS GroupName
    ,[Authen-Failure-Code] AS FailureCode
FROM
    [LogFiles]
WHERE
    UserName = '[UserName]'
    OR UserName LIKE '[UserName]@jmu.edu'
    OR UserName LIKE 'JMUAD\[UserName]'
ORDER BY Timestamp ASC
"@

        $logs = @()
        foreach ($c in $ComputerName) {
            $l = Get-ChildItem "\\$c\$LogPath" | Where-Object { $_.LastWriteTime -gt $Start } | Foreach-Object { $_.FullName }
            Write-Verbose "Found $($l.Count) log files on $c within the requested timeframe"
            $logs += $l
        }
        Write-Verbose "Will query $($logs.Count) total log files on $($ComputerName.Count) computer(s)"

        $logs = $logs | Foreach-Object { "`'$_`'" }
        $logs = $logs -join ",`n    "
        $query = $query.Replace("[LogFiles]", $logs)

        $messageMap = @{
            "Authen session timed out: Supplicant didnot respond to ACS correctly. Check supplicant configuration" = "Generic Client Config Problem"
            "EAP-TLS or PEAP authentication failed due to unknown CA certificate during SSL handshake" = "Client Certificate Trust Misconfiguration"
            "EAP-TLS or PEAP authentication failed during SSL handshake" = "Client Certificate Trust or Date/Time Misconfiguration"
            "EAP type not configured" = "Client EAP Misconfiguration; needs to select PEAP"
            "EAP_LEAP not configured or the NAS type defined does not support the protocol" = "Client EAP Misconfiguration; needs to select PEAP"
            "EAP_TLS Type not configured" = "Client EAP Misconfiguration; needs to select PEAP"
            "External DB user invalid or bad password" = "Invalid or bad password"
            "External DB account locked out" = "Account locked out"
        }
    }

    PROCESS {
        $fullQuery = $query.Replace("[UserName]", $User)
        #Write-Verbose "`n$fullQuery"

        $sw = New-Object System.Diagnostics.Stopwatch
        $sw.Start()
        $results = & "$LogParserPath" -O:CSV -stats:OFF "$fullQuery" | ConvertFrom-Csv
        $sw.Stop()
        Write-Verbose "Search ended (took $($sw.ElapsedMilliseconds)ms) and found $($results.Count) events"

        foreach ($r in $results) {
            $r.Timestamp = [DateTime]$r.Timestamp
            $r = Add-Member -InputObject $r -PassThru -MemberType NoteProperty -Name "ErrorMessage" -Value $null
            Add-Member -InputObject $r -MemberType NoteProperty -Name "Vendor" -Value $null

            if ($messageMap.Keys -contains $r.FailureCode) {
                Write-Verbose "Got here"
                $r.ErrorMessage = $messageMap[$r.FailureCode]
            }
        }

        $results
    }
}
