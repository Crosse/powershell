$hubs = Get-ExchangeServer | ? { $_.ServerRole -match 'Hub' }
$internalSend = New-Object System.Collections.ArrayList
$externalSend = New-Object System.Collections.ArrayList
$internalReceived = New-Object System.Collections.ArrayList
$externalReceived = New-Object System.Collections.ArrayList
$forwards = New-Object System.Collections.ArrayList

$hubs | Get-MessageTrackingLog -Start (Get-Date).AddDays(-1) -ResultSize Unlimited -Event | % {
    $log = $_
	if ($log.EventId -eq 'SEND') {
		if ($log.Sender -like "*@jmu.edu") {
			foreach ($recipient in $log.Recipients) {
				$fi = $false
				$fe = $false
				if ($recipient -like "*@jmu.edu" -and $fi -eq $false) {
					$internalSend.Add($log)
				}
				if ($recipient -notlike "*@jmu.edu" -and $fe -eq $false) {
					$externalSend.Add($log)
				}

				if ($fi -and $fe) { break }
			}
		} else {
			if ($log.Recipients -notmatch "@jmu.edu") {
				$forwards.Add($log)
			}
		}
	} elseif ($log.EventId -eq 'DELIVER') {
		if ($log.Sender -like "*@jmu.edu") {
			$internalReceived.Add($log)
		} else {
			$externalReceived.Add($log)
		}
	}
}


#$sentLogs | % { $totalSentSize += $_.TotalBytes }
#$receivedLogs | % { $totalReceiveSize += $_.TotalBytes }


#$avgSendSize = [Math]::Floor($totalSentSize / $sentLogs.Count / 1KB)
#$avgReceiveSize = [Math]::Floor($totalReceiveSize / $receivedLogs.Count / 1KB)

Write-Host "Total Sent Messages in the past 24 hours:  $($sentLogs.Count)"
Write-Host "Total Received Messages in the past 24 hours:  $($receivedLogs.Count)"
Write-Host "Average Sent Message Size:  $avgSendSize KB"
Write-Host "Average Received Message Size:  $avgReceiveSize KB"

