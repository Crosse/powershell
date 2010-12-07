param ($Server=$null) { }

if ([String]::IsNullOrEmpty($Server)) {
    $servers = Get-ExchangeServer | ? { $_.ServerRole -match "Hub" }
} else {
    $servers = @(Get-ExchangeServer $Server)
}

$entryObjs = New-Object System.Collections.ArrayList

foreach ($server in $servers) {
    Write-Host "Retrieving logs from $($server.Name)..."
    $logDir = "\\" + $server.name + "\"
    $logDir += (Get-TransportServer $server.Name).ConnectivityLogPath
    $logDir = $logDir.Replace("TransportRoles\Logs\Connectivity", "Logging\Imap4")
    $logDir = $logDir.Replace(":", "$")
    $logDir += "\*.log"

    $logLines = Get-Content $logDir

# Process each line into an object.
    $i = 1
    $count = $logLines.Count
    foreach ($line in $logLines) {
        if ($line.StartsWith("#")) { continue }
        if ([String]::IsNullOrEmpty($line)) { continue }

        $percent = $([int]($i/$count*100))
        if ($i % 250 -eq 0) {
            Write-Progress -Activity "Objectifying logs for $($server.Name)" `
                -Status "$percent% Complete" `
                -PercentComplete $percent -CurrentOperation "Processing line $i of $count..."
        }
        $i++

        $lineArray = $line.Split(',')

        $entry = New-Object PSObject
        $entry = Add-Member -PassThru -InputObject $entry NoteProperty Timestamp (Get-Date $lineArray[0])
        Add-Member -InputObject $entry NoteProperty SessionId $lineArray[2]
        Add-Member -InputObject $entry NoteProperty SequenceNumber $lineArray[3]
        Add-Member -InputObject $entry NoteProperty LocalEndpoint $lineArray[4]
        Add-Member -InputObject $entry NoteProperty RemoteEndpoint $lineArray[5]
        Add-Member -InputObject $entry NoteProperty Event $null
        Add-Member -InputObject $entry NoteProperty Data $lineArray[7]

        switch ($lineArray[6]) {
            "+" { $entry.Event = "Connect" }
            "-" { $entry.Event = "Disconnect" }
            ">" { $entry.Event = "Send" }
            "<" { $entry.Event = "Receive" }
            "*" { $entry.Event = "Information" }
        }

        $null = $entryObjs.Add($entry)
    }
    Write-Progress -Activity "Objectifying logs for $($server.Name)" `
        -Status "100% Complete" -PercentComplete 100 -Completed:$true
}

$entryObjs
