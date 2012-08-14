################################################################################
#
# $Id$
#
# DESCRIPTION:  Sends an email with relevant Exchange statistics to various
#               users.
#
# Copyright (c) 2009-2012 Seth Wright <wrightst@jmu.edu>
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
$To         = @("wrightst@jmu.edu")
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
[UInt64]$totalQuotaBytes = 0
$dbInfoArray = New-Object System.Collections.ArrayList
$userInfoArray = @{}
$statsArray = @{}
$dbStatsArray = @{}

Write-Verbose "Discovering all mailboxes..."
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
        $dbInfo = New-Object PSObject -Property @{
                        Identity            = $db.Name
                        Mailboxes           = $null
                        EdbFileSizeInGB     = $null
                        AvailableSpaceInMB  = $null
                        CommitPercent       = $null
                        BackupStatus        = $null
                        LastFullBackup      = $null
                        MountedOnServer     = $db.MountedOnServer.Split(".")[0]
                    }

        if ($db.DatabaseSize -ne $null) {
            $dbInfo.EdbFileSizeInGB = [Math]::Round($db.DatabaseSize.ToBytes()/1GB, 0)
            $totalStorageBytes += $db.DatabaseSize.ToBytes()
        }

        $dbInfo.LastFullBackup = $db.LastFullBackup
        if ($db.LastFullBackup -gt (Get-Date).AddDays(-1)) {
            $dbInfo.BackupStatus = "OK"
        } else {
            $dbInfo.BackupStatus = "NOT OK"
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
                    $userInfoArray[$user.sAMAccountName] = $stats.TotalItemSize.Value.ToBytes()

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
            $totalQuotaBytes += $totalDbUserQuota
        }

        $null = $dbInfoArray.Add($dbInfo)
        $i++
        #if ($i -gt 2) {
        #    break
        #}
    }
    Write-Progress -Activity "Gathering Database Statistics" -Status "Finished" -Id 1 -Completed
}

# Misnomer, really.  This gets the top senders by recipient count.
if ($IncludeTopRecipients) {
    Write-Verbose "Getting Top Recipient Counts"
    $recipientCounts = & "$cwd\Get-TopRecipientCounts.ps1" -Start $Start -Verbose:$Verbose
}

Write-Verbose "Getting UserMailbox count..."
$statsArray["User Mailboxes"] = @($allMailboxes | ? { $_.msExchRecipientTypeDetails -eq 1 }).Count
Write-Verbose "Found $($statsArray['User Mailboxes']) UserMailboxes"

Write-Verbose "Getting SharedMailbox count..."
$statsArray["Shared Mailboxes"] = @($allMailboxes | ? { $_.msExchRecipientTypeDetails -eq 4 }).Count
Write-Verbose "Found $($statsArray['Shared Mailboxes']) Shared Mailboxes"

Write-Verbose "Getting Resource Mailbox count..."
$statsArray["Resource Mailboxes"] = @($allMailboxes | ? { $_.msExchRecipientTypeDetails -eq 16 -or $_.msExchRecipientTypeDetails -eq 32 }).Count
Write-Verbose "Found $($statsArray['Resource Mailboxes']) Resource Mailboxes"

Write-Verbose "Getting Disabled Mailboxes count..."
$statsArray["Disabled Mailboxes"] = @($allMailboxes | ? { $_.msExchRecipientTypeDetails -eq 1 -and ($_.userAccountControl -band 2) -eq 2}).Count
Write-Verbose "Found $($statsArray['Disabled Mailboxes']) Disabled Mailboxes"

Write-Verbose "Getting MailUser count..."
$statsArray["Mail Users"] = (adfind -c -q -f '(&(!(homeMDB=*))(targetAddress=*)(objectClass=user))')[1].Split(" ")[0]
Write-Verbose "Found $($statsArray['Mail Users']) MailUsers"

Write-Verbose "Getting Distribution Group count..."
$statsArray["Distribution Groups"] = @(Get-DistributionGroup).Count
Write-Verbose "Found $($statsArray['Distribution Groups']) Distribution Groups"

if ($IncludeDatabaseStatistics) {
    $dbStatsArray["Storage Used (Databases)"] = "{0:N2} GB" -f ($totalStorageBytes/1GB)
    Write-Verbose "Storage Used (Databases):  $($dbStatsArray['Total Storage Used'])"

    $dbStatsArray["Storage Used (Mailbox)"] = "{0:N2} GB" -f (($userInfoArray.Values | Measure-Object -Sum).Sum/1GB)
    Write-Verbose "Storage Used (Mailboxes):  $($dbStatsArray['Actual Storage Used'])"

    $dbStatsArray["Total Quota Allocated"] = "{0:N2} GB" -f ($totalQuotaBytes/1GB)
    Write-Verbose "Total Quota Allocated:  $($dbStatsArray['Total Quota Allocated'])"
}

if ($IncludeMessageMetrics) {
    Write-Verbose "Getting Message Metrics"
    $hubs = Get-ExchangeServer | ? { $_.ServerRole -match 'Hub' }
    $messageStats = $hubs | Get-MessageTrackingLog -ResultSize Unlimited `
                                -Start $Start | ? {
                                    $_.Source -eq 'STOREDRIVER' -and
                                    ($_.EventId -eq 'DELIVER' -or
                                    $_.EventID -eq 'RECEIVE')
                                } | Measure-Object -Sum -Property TotalBytes

    # IOPS calculations taken from here:
    # http://technet.microsoft.com/en-us/library/ee832791.aspx
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

$bodyStyle          = "font-family:sans-serif;color:black;background-color:white;"
$tableBaseStyle     = "border:1px solid green;border-collapse:collapse;margin:25px;margin-botom:50px;float:left;"
$tableStyle         = "$tableBaseStyle;width:350px;"
$tableDatabaseStyle = "$tableBaseStyle;width:750px;"
$captionStyle       = "border-style:none;border-collapse:collapse;text-align:center;font-weight:bold;font-size:1.2em;background-color:white;margin:4px;"
$thStyle            = "border:1px solid green;border-collapse:collapse;text-align:center;margin:4px;color:#333333;"
$trStyle            = "margin:4px;"
$trAltStyle         = "margin:4px;background-color:#EAF2D3"
$tdStyle            = "margin:4px;border:1px solid green;border-collapse:collapse;text-align:right;"
$tdIdentityStyle    = "margin:4px;border:1px solid green;border-collapse:collapse;text-align:left;"
$tdServerStyle      = "margin:4px;border:1px solid green;border-collapse:collapse;text-align:center;"
$tdWarningStyle     = "margin:4px;border:1px solid green;border-collapse:collapse;background-color:yellow;"
$tdErrorStyle       = "margin:4px;border:1px solid green;border-collapse:collapse;background-color:red;"
$brClearStyle       = "clear:both;"

$Body  = @"
<!DOCTYPE html PUBLIC "-//W3C/DTD XHTML 1.0 Tranisional//EN"
    "http://www.w3.org/xhtml1/DTD/xhtml1-transitional.dtd">
<html>
    <head>
        <title>$Title</title>
    </head>
    <body style="$bodyStyle">
    <h1>Exchange Statistics for $(Get-Date -Format d)</h1>

    <table style="$tableStyle">
        <caption style="$captionStyle">User Statistics</caption>
        <tr style="$trStyle">
            <th style="$thStyle">Metric</th>
            <th style="$thStyle">Value</th>
        </tr>

"@

$i = 0
foreach ($key in $statsArray.Keys) {
    if (($i % 2) -eq 0) {
        $style = $trAltStyle
    } else {
        $style = $tr
    }

    $Body += @"
        <tr style="$style">
            <td style="$tdStyle">$($key)</td>
            <td style="$tdStyle">$($statsArray[$key])</td>
        </tr>

"@
    $i++
}

foreach ($key in $dbStatsArray.Keys) {
    if (($i % 2) -eq 0) {
        $style = $trAltStyle
    } else {
        $style = $tr
    }

    $Body += @"
        <tr style="$style">
            <td style="$tdStyle">$($key)</td>
            <td style="$tdStyle">$($dbStatsArray[$key])</td>
        </tr>

"@
    $i++
}

$Body += @"
    </table>

"@

if ($IncludeMessageMetrics) {
    $Body += @"
    <table style="$tableStyle">
        <caption style="$captionStyle">Message Statistics</caption>
        <tr style="$trStyle">
            <th style="$thStyle">Metric</th>
            <th style="$thStyle">Value</th>
        </tr>
        <tr style="$trAltStyle">
            <td style="$tdStyle">Average Message Size</td>
            <td style="$tdStyle">$([Math]::Round($avgMessageSizeInKB, 0)) KB</td>
        </tr>
        <tr style="$trStyle">
            <td style="$tdStyle">Average Messages Sent/Rec'd per Mailbox</td>
            <td style="$tdStyle">$([Math]::Round($avgMessagesPerMailbox, 0))</td>
        </tr>
        <tr style="$trAltStyle">
            <td style="$tdStyle">Predicted Total IOPS per Mailbox</td>
            <td style="$tdStyle">$([Math]::Round($iopsPerMailbox, 2))</td>
        </tr>
        <tr style="$trStyle">
            <td style="$tdStyle">Predicted Total IOPS</td>
            <td style="$tdStyle">$([Math]::Round($totalIops, 2))</td>
        </tr>
    </table>

"@
}

$Body += @"
    <br style="$brClearStyle" />

"@

if ($IncludeTopRecipients) {
    $Body += @"
    <table style="$tableStyle">
        <caption style="$captionStyle">Top Senders by Recipient Count</caption>
        <tr style="$trStyle">
            <th style="$thStyle">Sender</th>
            <th style="$thStyle">Recipients</th>
        </tr>

"@
    for ($i = 0; $i -lt $recipientCounts.Count; $i++) {
        if (($i % 2) -eq 0) {
            $style = $trAltStyle
        } else {
            $style = $tr
        }

        $Body += @"
        <tr style="$style">
            <td style="$tdIdentityStyle">$($recipientCounts[$i].Key)</td>
            <td style="$tdStyle">$($recipientCounts[$i].Value)</td>
        </tr>

"@
    }
    $Body += @"
    </table>

"@
}

if ($IncludeTopStorageUsers) {
    $Body += @"
    <table style="$tableStyle">
        <caption style="$captionStyle">Top Users by Mailbox Size</caption>
        <tr style="$trStyle">
            <th style="$thStyle">User</th>
            <th style="$thStyle">Mailbox Size</th>
        </tr>

"@

    $i = 0
    foreach ($user in ($userInfoArray.GetEnumerator() | Sort Value -Descending | Select -First 10)) {
        if (($i % 2) -eq 0) {
            $style = $trAltStyle
        } else {
            $style = $tr
        }

        $Body += @"
        <tr style="$style">
            <td style="$tdIdentityStyle">$($user.Name)</td>
            <td style="$tdStyle">$([Math]::Round($user.Value/1MB, 0)) MB</td>
        </tr>

"@
        $i++
    }

    $Body += @"
    </table>

"@

}
$Body += @"
    <br style="$brClearStyle" />

"@


if ($IncludeDatabaseStatistics) {
    $Body += @"
        <table style="$tableDatabaseStyle">
            <caption style="$captionStyle">Database Information</caption>
            <tr style="$trStyle">
                <th style="$thStyle">Database</th>
                <th style="$thStyle">Mailbox Count</th>
                <th style="$thStyle">EDB File Size</th>
                <th style="$thStyle">Available Space</th>
                <th style="$thStyle">Commit Percentage</th>
                <th style="$thStyle">Last Full Backup</th>
                <th style="$thStyle">Backup Status</th>
                <th style="$thStyle">Mounted On Server</th>
            </tr>

"@

    foreach ($db in $dbInfoArray) {
        $Body += @"
            <tr style="$trStyle">
                <td style="$tdIdentityStyle">$($db.Identity)</td>
                <td style="$tdStyle">$($db.Mailboxes)</td>
                <td style="$tdStyle">$($db.EdbFileSizeInGB) GB</td>
                <td style="$tdStyle">$($db.AvailableSpaceInMB) MB</td>

"@
        if ($db.CommitPercent -eq "unlimited") {
            $Body += "<td style=`"$tdStyle`">{0}</td>" -f $db.CommitPercent
        } elseif ($db.CommitPercent -gt 100) {
            $Body += "<td style=`"$tdWarningStyle`">{0:N0}%</td>" -f $db.CommitPercent
        } elseif ($db.CommitPercent -gt 150) {
            $Body += "<td style=`"$tdErrorStyle`">{0:N0}%</td>" -f $db.CommitPercent
        } else {
            $Body += "<td style=`"$tdStyle`">{0:N0}%</td>" -f $db.CommitPercent
        }

        $Body += "`n<td style=`"$tdStyle`">$($db.LastFullBackup.ToString('M/d/yy HH:mm'))</td>`n"

        if ($db.BackupStatus -match "NOT OK") {
            $Body += "<td style=`"$tdWarningStyle`">{0}</td>" -f $($db.BackupStatus)
        } else {
            $Body += "<td style=`"$tdStyle`">{0}</td>" -f $($db.BackupStatus)
        }

        $Body += @"

                <td style="$tdServerStyle">$($db.MountedOnServer.ToLower())</td>
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
