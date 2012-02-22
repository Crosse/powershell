################################################################################
#
# $Id$
#
# DESCRIPTION:  Sends an email with relevant Exchange statistics to various
#               users.
#
# Copyright (c) 2009,2010 Seth Wright <wrightst@jmu.edu>
#
# Permission to use, copy, modify, and distribute this software for any
# purpose with or without fee is hereby granted, provided that the above
# copyright notice and this permission notice appear in all copies.
#
# THE SOFTWARE IS PROVIDED "AS IS" AND THE AUTHOR DISCLAIMS ALL WARRANTIES
# WITH REGARD TO THIS SOFTWARE INCLUDING ALL IMPLIED WARRANTIES OF
# MERCHANTABILITY AND FITNESS. IN NO EVENT SHALL THE AUTHOR BE LIABLE FOR
# ANY SPECIAL, DIRECT, INDIRECT, OR CONSEQUENTIAL DAMAGES OR ANY DAMAGES
# WHATSOEVER RESULTING FROM LOSS OF USE, DATA OR PROFITS, WHETHER IN AN
# ACTION OF CONTRACT, NEGLIGENCE OR OTHER TORTIOUS ACTION, ARISING OUT OF
# OR IN CONNECTION WITH THE USE OR PERFORMANCE OF THIS SOFTWARE.
#
################################################################################

[CmdletBinding(SupportsShouldProcess=$true,
        ConfirmImpact="High")]

param (
        [switch]
        $IncludeDatabaseStatistics=$true,

        [switch]
        $IncludeTopRecipients=$true,

        [switch]
        $IncludeTopStorageUsers=$true,

        [switch]
        $IncludeMessageMetrics=$true,

        [switch]
        $RunSetMailboxQuotaLimits=$false,

        [DateTime]
        $Start=(Get-Date).AddDays(-1)
      )

# Change these to suit your environment
$SmtpServer = "mailgw.jmu.edu"
$From       = "it-exmaint@jmu.edu"
#$To         = @("wrightst@jmu.edu", "gumgs@jmu.edu", "liskeygn@jmu.edu", "stockntl@jmu.edu", "najdziav@jmu.edu")
$To         = "wrightst@jmu.edu"
$Title      = "Exchange Statistics for $(Get-Date -Format d)"
$MaxDatabaseSizeInBytes = 250*1GB

##################################
$cwd = [System.IO.Path]::GetDirectoryName(($MyInvocation.MyCommand).Definition)

Write-Verbose "SmtpServer: $SmtpServer"
Write-Verbose "From: $From"
Write-Verbose "To: $To"
Write-Verbose "Email Subject:  $Title"
Write-Verbose "MaxDatabaseSizeInBytes:  $($MaxDatabaseSizeInBytes/1GB)GB"
Write-Verbose "Start Date: $Start"
Write-Verbose "Script path:  $cwd"

[UInt64]$totalStorageBytes = 0
$dbInfoArray = New-Object System.Collections.ArrayList
$userInfoArray = @{}
$statsArray = @{}

Write-Host "Discovering all mailboxes..."
$allMailboxes = $(adfind -csv -q -f '(&(!(cn=SystemMailbox*))(homeMDB=*)(objectClass=user))' sAMAccountName homeMDB msExchRecipientTypeDetails userAccountControl mDBUseDefaults | ConvertFrom-Csv)
Write-Verbose "Found $($allMailboxes.Count) mailboxes"

if ($IncludeDatabaseStatistics) {
    $dbs = @(Get-MailboxDatabase -Status | Sort whenCreated)
    $dbCount = $dbs.Count
    $i = 1
    $dstart = Get-Date
    foreach ($db in $dbs) {
        $dend = Get-Date
        $dtotalSeconds = ($dend - $dstart).TotalSeconds
        $timePerDb = $dtotalSeconds / $i
        $dtimeLeft = $timePerDb * ($dbCount - $i)

        Write-Progress  -Activity "Gathering Database Statistics" `
                        -Status $db.Name `
                        -PercentComplete ($i/$dbCount*100) `
                        -Id 1 -SecondsRemaining $dtimeLeft
        $dbInfo = New-Object PSObject
        $dbInfo = Add-Member -PassThru -InputObject $dbInfo NoteProperty Identity $db.Name
        Add-Member -InputObject $dbInfo NoteProperty Mailboxes $null
        Add-Member -InputObject $dbInfo NoteProperty EdbFileSizeInGB $null
        Add-Member -InputObject $dbInfo NoteProperty AvailableSpaceInMB $null
        Add-Member -InputObject $dbInfo NoteProperty CommitPercent $null
        Add-Member -InputObject $dbInfo NoteProperty BackupStatus $null
        Add-Member -InputObject $dbInfo NoteProperty LastFullBackup $null
        Add-Member -InputObject $dbInfo NoteProperty MountedOnServer $db.MountedOnServer.Split(".")[0]

        if ($db.DatabaseSize -ne $null) {
            $dbInfo.EdbFileSizeInGB = [Math]::Round($db.DatabaseSize.ToBytes()/1GB, 0)
            $totalStorageBytes += $db.DatabaseSize.ToBytes()
        }

        $dbInfo.LastFullBackup = $db.LastFullBackup

        if ($db.LastFullBackup -gt (Get-Date).AddDays(-1)) {
            $dbInfo.BackupStatus = "OK (<24h)"
        } else {
            $dbInfo.BackupStatus = "NOT OK (>24h)"
        }

        $dbInfo.AvailableSpaceInMB = [Math]::Round($db.AvailableNewMailboxSpace.ToBytes()/1MB, 0)
        $dbUsers = @($allMailboxes | ? { $_.homeMDB -match "CN=$($db.Name)," })
        $dbInfo.Mailboxes = $dbUsers.Count

        if ($db.ProhibitSendReceiveQuota -eq "unlimited") {
            Write-Verbose "Database quota set to unlimited"
        } else {
            [UInt64]$totalDbUserQuota = 0
            $dbQuota = $db.ProhibitSendReceiveQuota.Value.ToBytes()
        }

        $usersCount = $dbUsers.Count
        $j = 0
        $startTime = Get-Date
        foreach ($user in $dbUsers) {
            if ($db.ProhibitSendReceiveQuota -ne "unlimited") {
                if ($user.mDBUseDefaults -eq $true) {
                    $totalDbUserQuota += $dbQuota
                } else {
                    $userQuota = (Get-Mailbox $user.sAMAccountName).ProhibitSendReceiveQuota
                    if ($userQuota.IsUnlimited -eq $false) {
                        $totalDbUserQuota += $userQuota.Value.ToBytes()
                    }
                }
            }

            if ($IncludeTopStorageUsers -or $RunSetMailboxQuotaLimits) {
                $stats = Get-MailboxStatistics $user.sAMAccountName
                if ($stats -ne $null) {
                    if ($IncludeTopRecipients) {
                        $userInfoArray[$user.sAMAccountName] = $stats.TotalItemSize.Value.ToBytes()
                    }

                    if ($RunSetMailboxQuotaLimits -and
                            (Test-Path "$cwd\Set-MailboxQuotaLimits.ps1") -and
                            ($stats.StorageLimitStatus -eq 'IssueWarning' -or
                            $stats.StorageLimitStatus -eq 'ProhibitSend' -or
                            $stats.StorageLimitStatus -eq 'MailboxDisabled')) {
                        Write-Verbose "Raising quota for user $($user.sAMAccountName)"
                    & "$cwd\Set-MailboxQuotaLimits.ps1" -Identity $user.sAMAccountName -Verbose:$Verbose -Confirm:$false
                    }
                }
            }
            $j++
            $end = Get-Date
            $totalSeconds = ($end - $startTime).TotalSeconds
            $timePerUser = $totalSeconds / $j
            $timeLeft = $timePerUser * ($usersCount - $j)
            Write-Progress  -Activity "Gathering User Statistics for Database $($db.Name)" `
                            -Status $user.sAMAccountName `
                            -PercentComplete ($j/$usersCount*100) `
                            -Id 2 -ParentId 1 `
                            -SecondsRemaining $timeLeft
        }
        Write-Progress -Activity "Gathering User Statistics" -Status "Finished" -Id 2 -ParentId 1 -Completed

        if ($db.ProhibitSendReceiveQuota -eq "unlimited") {
            $dbInfo.CommitPercent = "unlimited"
        } else {
            $dbInfo.CommitPercent = [Math]::Round(($totalDbUserQuota/$MaxDatabaseSizeInBytes*100), 0)
        }

        Write-Verbose "Database Name: $($dbInfo.Identity)"
        Write-Verbose "Database Size: $($dbInfo.EdbFileSizeInGB)"
        Write-Verbose "Database Available Space: $($dbInfo.AvailableSpaceInMB)MB"
        Write-Verbose "Database Commit %: $($dbInfo.CommitPercent)"
        Write-Verbose "Database Last Full Backup: $($dbInfo.LastFullBackup)"
        Write-Verbose "Database Backup Status: $($dbInfo.BackupStatus)"
        Write-Verbose "Database Mounted on:  $($dbInfo.MountedOnServer)"
        $null = $dbInfoArray.Add($dbInfo)
        $i++
        #if ($i -gt 2) {
        #    break
        #}
    }
    Write-Progress -Activity "Gathering Database Statistics" -Status "Finished" -Id 1 -Completed
}

if ($IncludeTopRecipients) {
    Write-Host "Getting Top Recipient Counts"
    $recipientCounts = & "$cwd\Get-TopRecipientCounts.ps1" -Start $Start -Verbose:$Verbose
}

Write-Host "Getting UserMailbox count..."
$statsArray["User Mailboxes"] = @($allMailboxes | ? { $_.msExchRecipientTypeDetails -eq 1 }).Count
Write-Verbose "Found $($statsArray['User Mailboxes']) UserMailboxes"

Write-Host "Getting SharedMailbox count..."
$statsArray["Shared Mailboxes"] = @($allMailboxes | ? { $_.msExchRecipientTypeDetails -eq 4 }).Count
Write-Verbose "Found $($statsArray['Shared Mailboxes']) Shared Mailboxes"

Write-Host "Getting Resource Mailbox count..."
$statsArray["Resource Mailboxes"] = @($allMailboxes | ? { $_.msExchRecipientTypeDetails -eq 16 -or $_.msExchRecipientTypeDetails -eq 32 }).Count
Write-Verbose "Found $($statsArray['Resource Mailboxes']) Resource Mailboxes"

Write-Host "Getting Disabled Mailboxes count..."
$statsArray["Disabled Mailboxes"] = @($allMailboxes | ? { $_.msExchRecipientTypeDetails -eq 1 -and ($_.userAccountControl -band 2) -eq 2}).Count
Write-Verbose "Found $($statsArray['Disabled Mailboxes']) Disabled Mailboxes"

Write-Host "Getting MailUser count..."
$statsArray["Mail Users"] = (adfind -c -q -f '(&(!(homeMDB=*))(targetAddress=*)(objectClass=user))')[1].Split(" ")[0]
Write-Verbose "Found $($statsArray['Mail Users']) MailUsers"

Write-Host "Getting Distribution Group count..."
$statsArray["Distribution Groups"] = @(Get-DistributionGroup).Count
Write-Verbose "Found $($statsArray['Distribution Groups']) Distribution Groups"

if ($IncludeDatabaseStatistics) {
    $statsArray["Total Storage Used"] = "{0:N2} GB" -f ($totalStorageBytes/1GB)
    Write-Verbose "Total Storage Used:  $($statsArray['Total Storage Used'])"
}

if ($IncludeMessageMetrics) {
    Write-Host "Getting Message Metrics"
    $hubs = Get-ExchangeServer | ? { $_.ServerRole -match 'Hub' }
    $messageStats = $hubs | Get-MessageTrackingLog -ResultSize Unlimited `
                                -Start $Start | ? {
                                    $_.Source -eq 'STOREDRIVER' -and
                                    ($_.EventId -eq 'DELIVER' -or
                                    $_.EventID -eq 'RECEIVE')
                                } | Measure-Object -Sum -Property TotalBytes

    $totalMessages = $messageStats.Count
    $avgMessageSizeInKB = ($messageStats.Sum / $totalMessages)/1KB
    $avgMessagesPerMailbox = $totalMessages / $allMailboxes.Count
    $iopsPerMailbox = ($avgMessagesPerMailbox / 1000) * ($avgMessageSizeInKB/75)
    $totalIops = $iopsPerMailbox * $allMailboxes.Count

    Write-Verbose "Total messages, last 24h:  $totalMessages"
    Write-Verbose "Average message size:  $avgMessageSizeInKB KB"
    Write-Verbose "Average messages per mailbox: $avgMessagesPerMailbox"
    Write-Verbose "Predicted IOPS per mailbox: $iopsPerMailbox"
    Write-Verbose "Total Predicted IOPS:  $totalIops"
}

$Body  = @"
<!DOCTYPE html PUBLIC "-//W3C/DTD XHTML 1.0 Tranisional//EN"
    "http://www.w3.org/xhtml1/DTD/xhtml1-transitional.dtd">
<html>
    <head>
        <style type="text/css">
            body {
                font-family:sans-serif;
                color:black;
                background-color:white;
            }
            table {
                border:1px solid green;
                border-collapse:collapse;
                margin:25px;
                margin-botom:50px;
                width:350px;
                float:left;
            }
            table.database {
                width:750px;
            }
            caption {
                border-style:none;
                border-collapse:collapse;
                text-align:center;
                font-weight:bold;
                font-size:1.2em;
                background-color:white;
                margin:4px;
            }
            th {
                border:1px solid green;
                border-collapse:collapse;
                text-align:center;
                margin:4px;
                /* background-color:#A7C942; */
                color:#333333;
            }
            tr {
                margin:4px;
            }
            tr.alt {
                background-color:#EAF2D3;
                margin:4px;
            }
            td {
                border:1px solid green;
                border-collapse:collapse;
                text-align:right;
                margin:4px;
            }
            td.identity {
                margin:4px;
                text-align:left;
            }
            td.server {
                margin:4px;
                text-align:center;
            }
            td.warning {
                margin:4px;
                background-color:yellow;
            }
            td.error {
                margin:4px;
                background-color:red;
            }
            br.clear {
                clear:both;
            }
        </style>
    </head>
    <body>

    <h1>Exchange Statistics for $(Get-Date -Format d)</h1>

    <table>
        <caption>User Statistics</caption>
        <tr>
            <th>Metric</th>
            <th>Value</th>
        </tr>

"@

$i = 0
foreach ($key in $statsArray.Keys) {
    if (($i % 2) -eq 0) {
        $alt = 'class="alt"'
    } else {
        $alt = ""
    }

    $Body += @"
        <tr $alt>
            <td>$($key)</td>
            <td>$($statsArray[$key])</td>
        </tr>

"@
    $i++
}

$Body += @"
    </table>

"@

if ($IncludeMessageMetrics) {
    $Body += @"
    <table>
        <caption>Message Statistics</caption>
        <tr>
            <th>Metric</th>
            <th>Value</th>
        </tr>
        <tr class="alt">
            <td>Average Message Size</td>
            <td>$([Math]::Round($avgMessageSizeInKB, 0)) KB</td>
        </tr>
        <tr>
            <td>Average Messages Sent/Rec'd per Mailbox</td>
            <td>$([Math]::Round($avgMessagesPerMailbox, 0))</td>
        </tr>
        <tr class="alt">
            <td>Predicted Total IOPS per Mailbox</td>
            <td>$([Math]::Round($iopsPerMailbox, 2))</td>
        </tr>
        <tr>
            <td>Predicted Total IOPS</td>
            <td>$([Math]::Round($totalIops, 2))</td>
        </tr>
    </table>

"@
}

$Body += @"
    <br class="clear"/>

"@

if ($IncludeTopRecipients) {
    $Body += @"
    <table>
        <caption>Top Senders by Recipient Count</caption>
        <tr>
            <th>Sender</th>
            <th>Recipients</th>
        </tr>

"@
    for ($i = 0; $i -lt $recipientCounts.Count; $i++) {
        if (($i % 2) -eq 0) {
            $alt = 'class="alt"'
        } else {
            $alt = ""
        }

        $Body += @"
        <tr $alt>
            <td class="identity">$($recipientCounts[$i].Key)</td>
            <td>$($recipientCounts[$i].Value)</td>
        </tr>

"@
    }
    $Body += @"
    </table>

"@
}

if ($IncludeTopStorageUsers) {
    $Body += @"
    <table>
        <caption>Top Users by Mailbox Size</caption>
        <tr>
            <th>User</th>
            <th>Mailbox Size</th>
        </tr>

"@

    $i = 0
    foreach ($user in ($userInfoArray.GetEnumerator() | Sort Value -Descending | Select -First 10)) {
        if (($i % 2) -eq 0) {
            $alt = 'class="alt"'
        } else {
            $alt = ""
        }

        $Body += @"
        <tr $alt>
            <td class="identity">$($user.Name)</td>
            <td>$([Math]::Round($user.Value/1MB, 0)) MB</td>
        </tr>

"@
        $i++
    }

    $Body += @"
    </table>

"@

}
$Body += @"
    <br class="clear"/>

"@


if ($IncludeDatabaseStatistics) {
    $Body += @"
        <table class="database">
            <caption>Database Information</caption>
            <tr>
                <th>Database</th>
                <th>Mailbox Count</th>
                <th>EDB File Size</th>
                <th>Available Space</th>
                <th>Commit Percentage</th>
                <th>Last Full Backup</th>
                <th>Backup Status</th>
                <th>Mounted On Server</th>
            </tr>

"@

    foreach ($db in $dbInfoArray) {
        $Body += @"
            <tr>
                <td class="identity">$($db.Identity)</td>
                <td>$($db.Mailboxes)</td>
                <td>$($db.EdbFileSizeInGB) GB</td>
                <td>$($db.AvailableSpaceInMB) MB</td>

"@
        if ($db.CommitPercent -gt 100) {
            $Body += '<td class="warning">{0:N0}%</td>' -f $db.CommitPercent
        } elseif ($db.CommitPercent -gt 150) {
            $Body += '<td class="error">{0:N0}%</td>' -f $db.CommitPercent
        } else {
            $Body += '<td>{0:N0}%</td>' -f $db.CommitPercent
        }

        $Body += "`n<td>$($db.LastFullBackup.ToString('M/d/yy HH:mm'))</td>`n"

        if ($db.BackupStatus -match 'NOT OK') {
            $Body += '<td class="warning">{0}</td>' -f $($db.BackupStatus)
        } else {
            $Body += '<td>{0}</td>' -f $($db.BackupStatus)
        }

        $Body += @"

                <td class="server">$($db.MountedOnServer.ToLower())</td>
            </tr>

"@
    }
    $Body += @"
    </table>
    <br />

"@
}

$Body += @"
    </body>
</html>
"@

Send-MailMessage -From $From -To $To -Subject $Title -Body $Body -SmtpServer $SmtpServer -UseSsl -BodyAsHtml
Write-Verbose "Sent Email"
