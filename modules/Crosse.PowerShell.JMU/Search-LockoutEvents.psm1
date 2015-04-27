function Search-LockoutEvents {
    [CmdletBinding(SupportsShouldProcess=$true, ConfirmImpact="High")]

    param (
            [Parameter(Mandatory=$true, ValueFromPipeline=$true)]
            [String]
            # The user to search for.
            $User,

            [Parameter(Mandatory=$false)]
            [DateTime]
            # The starting time for the search.  The default is one hour from the current time.
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
            $ResultSize = [Long]::MaxValue
          )

    $sw = New-Object System.Diagnostics.Stopwatch
    $startTime = $Start.ToUniversalTime().ToString('yyyy-MM-ddTHH:mm:ss.fffZ')
    $endTime = $End.ToUniversalTime().ToString('yyyy-MM-ddTHH:mm:ss.fffZ')
    $xPathQuery = @"
<QueryList>
<Query Id="0">
    <Select Path="$EventLogName">
        *[System[TimeCreated[@SystemTime &gt;= "$startTime" and @SystemTime &lt;= "$endTime"]]] 
        and 
        *[System[(EventID=4740 or EventID=4625)]]
        and 
        *[EventData[Data[@Name="TargetUserName"] = "$User" ]]
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
        if ($ResultSize) {
            $events += Get-WinEvent -ComputerName $c `
                                    -LogName $EventLogName `
                                    -FilterXPath $xPathQuery `
                                    -MaxEvents $ResultSize `
                                    -ErrorAction SilentlyContinue
        } else {
            $events += Get-WinEvent -ComputerName $c `
                                    -LogName $EventLogName `
                                    -FilterXPath $xPathQuery #`
                                    #-ErrorAction SilentlyContinue
        }
        $sw.Stop()
        Write-Verbose "Search ended on $c (took $($sw.ElapsedMilliseconds)ms) and found $($events.Count) events"
        foreach ($event in $events) {
            $props = Get-EventData -Event $event -Property TargetUserName, SubjectUserName, IpAddress, WorkstationName
            New-Object PSObject -Property @{
                UserName = $props["TargetUserName"]
                ServerName = $props["SubjectUserName"]
                WorkstationName = $props["WorkstationName"]
                IPAddress = $props["IpAddress"]
            }
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
    $root = $xml.DocumentElement
    $nsmgr = New-Object System.Xml.XmlNamespaceManager $xml.NameTable
    $nsmgr.AddNamespace("event", $xml.DocumentElement.NamespaceURI)
    $values = @{}

    if ($Property.Count -eq 0) {
        Write-Verbose "No explictly-requestes properties; returning all"
        foreach ($prop in $root.SelectNodes("//event:Data", $nsmgr)) {
            $Property += $prop.Name
        }
    }

    foreach ($prop in $Property) {
        $query = "//event:Data[@Name='$prop']/text()"
        $values[$prop] = $root.SelectNodes($query, $nsmgr).Data
    }
    $values
}
