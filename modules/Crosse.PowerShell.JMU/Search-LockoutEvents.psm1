function Search-LockoutEvents {
    [CmdletBinding(SupportsShouldProcess=$true, ConfirmImpact="High")]

    param (
            [Parameter(Mandatory=$false, ValueFromPipeline=$true)]
            [String]
            # The user to search for.  If no user is specified, all lockout events will be returned.
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
            [String]
            # The event log to search for lockout events.
            $EventLogName = "Security",

            [Parameter(Mandatory=$false)]
            # The number of events to return.  By default, all events are returned.
            [Long]
            $ResultSize = [Long]::MaxValue,

            [Parameter(Mandatory=$false)]
            [Switch]
            $FuzzySearch = $false
          )

    BEGIN {
        $processLookup = @{
            'EdgeTransport.exe' = 'SMTP'
            'Microsoft.Exchange.Imap4.exe' = 'IMAP'
            'w3wp.exe' = 'OWA/Phone'
            'Microsoft.Exchange.Pop3.exe' = 'POP'
        }

        if ($FuzzySearch) {
            Write-Verbose "Fuzzy searching has been enabled.  This will take longer than a strict search to perform, but may also return more (better?) results."
        }
    }

    PROCESS {
        $sw = New-Object System.Diagnostics.Stopwatch
        $startTime = $Start.ToUniversalTime().ToString('yyyy-MM-ddTHH:mm:ss.fffZ')
        $endTime = $End.ToUniversalTime().ToString('yyyy-MM-ddTHH:mm:ss.fffZ')

        if (![String]::IsNullOrEmpty($User) -and !$FuzzySearch) {
            $userQuery = @"
        and *[EventData[Data[@Name="TargetUserName"] = "$User" ]]
"@
        }

        $xPathQuery = @"
<QueryList>
<Query Id="0" Path="$EventLogName">
    <Select Path="$EventLogName">
        *[System[(EventID=4740 or EventID=4625) and TimeCreated[@SystemTime &gt;= "$startTime" and @SystemTime &lt;= "$endTime"]]]
$userQuery
    </Select>
</Query>
</QueryList>
"@

        $events = @()
        $total = $ComputerName.Count
        foreach ($c in $ComputerName) {
            $pctComplete = ($ComputerName.Count - $total) / $ComputerName.Count * 100
            $total--
            Write-Progress -Activity "Searching event logs" -Status "Querying $c" -PercentComplete $pctComplete
            Write-Verbose "Starting search on $c"
            $sw.Start()
            $session = New-Object System.Diagnostics.Eventing.Reader.EventLogSession $c
            $logQuery = New-Object System.Diagnostics.Eventing.Reader.EventLogQuery $EventLogName, "LogName", $xPathQuery
            $logQuery.Session = $session
            $logReader = New-Object System.Diagnostics.Eventing.Reader.EventLogReader $logQuery

            $i = 0
            while (($evt = $logReader.ReadEvent()) -ne $null) {
                $props = Get-EventData -Event $evt

                if ($FuzzySearch -and $props["TargetUserName"] -notlike "*$User*") {
                    continue
                }

                $source = ""
                if ($props["ProcessName"]) {
                    foreach ($p in $processLookup.Keys) {
                        if ($props["ProcessName"].EndsWith($p)) {
                            $source = $processLookup[$p]
                            break
                        }
                    }
                }
                switch ($evt.Id) {
                    4740 { $eventType = "AccountLockout" }
                    4625 { $eventType = "InvalidPasswordAttempt" }
                    default { $eventType = "SethMessedUp" }
                }
                New-Object PSObject -Property @{
                    RecordId = $evt.RecordId
                    TimeCreated = $evt.TimeCreated
                    UserName = $props["TargetUserName"]
                    TargetDomainName = $props["TargetDomainName"]
                    WorkstationName = $props["WorkstationName"]
                    IPAddress = $props["IpAddress"]
                    ProcessName = $props["ProcessName"]
                    Source = $source
                    EventType = $eventType
                }
                $i++
                if ($i -ge $ResultSize) { break }
            }
            $sw.Stop()
            Write-Verbose "Search ended on $c (took $($sw.ElapsedMilliseconds)ms) and found $i events"
        }
        Write-Progress -Activity "Searching event logs" -Status "Done." -Completed
    }
}

function Get-EventData {
    [CmdletBinding()]
    param (
            [Parameter(Mandatory=$true,
                ValueFromPipeline=$true)]
            [ValidateNotNullOrEmpty()]
            [System.Diagnostics.Eventing.Reader.EventLogRecord]
            $Event,

            [Parameter(Mandatory=$false)]
            [String[]]
            $Property = @()
          )

    $xml = [xml]$Event.ToXml()
    $nsmgr = New-Object System.Xml.XmlNamespaceManager $xml.NameTable
    $nsmgr.AddNamespace("event", $xml.DocumentElement.NamespaceURI)

    $values = @{}
    $root = $xml.DocumentElement
    if ($Property.Count -eq 0) {
        foreach ($prop in $root.SelectNodes("//event:Data", $nsmgr)) {
            $Property += $prop.Name
        }
    }

    foreach ($prop in $Property) {
        $query = "//event:Data[@Name='$prop']/text()"
        $values[$prop] = $root.SelectSingleNode($query, $nsmgr).Data
    }
    $values
}
