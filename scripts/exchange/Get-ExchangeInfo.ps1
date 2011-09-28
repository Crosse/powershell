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
        $IncludeTopRecipients=$true,

        [switch]
        $IncludeTopStorageUsers=$true,

        [switch]
        $RunSetMailboxQuotaLimits=$false
      )

# Change these to suit your environment
$SmtpServer = "mailgw.jmu.edu"
$From       = "it-exmaint@jmu.edu"
#$To         = "wrightst@jmu.edu, gumgs@jmu.edu, liskeygn@jmu.edu, stockntl@jmu.edu, flynngn@jmu.edu, najdziav@jmu.edu"
$To         = "wrightst@jmu.edu"
$Title      = "Exchange User Detail for $(Get-Date -Format d)"
$MaxDatabaseSizeInBytes = 250*1GB

##################################
$cwd = [System.IO.Path]::GetDirectoryName(($MyInvocation.MyCommand).Definition)

Write-Verbose "SmtpServer: $SmtpServer"
Write-Verbose "From: $From"
Write-Verbose "To: $To"
Write-Verbose "Email Subject:  $Title"
Write-Verbose "MaxDatabaseSizeInBytes:  $($MaxDatabaseSizeInBytes/1GB)GB"
Write-Verbose "Script path:  $cwd"

[UInt64]$totalStorageBytes = 0
$dbInfoArray = New-Object System.Collections.ArrayList
$userInfoArray = @{}
$statsArray = @{}

Write-Host "Discovering all mailboxes..."
$allMailboxes = $(adfind -csv -q -f '(&(!(cn=SystemMailbox*))(homeMDB=*))' cn homeMDB msExchRecipientTypeDetails userAccountControl mDBUseDefaults | ConvertFrom-Csv)
Write-Verbose "Found $($allMailboxes.Count) mailboxes"

# This foreach clause *should* look weird, but it's the best way to ensure 
# that the mailbox databases are sorted by number.
$dbs = Get-MailboxDatabase -Status | Sort Name
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
    Add-Member -InputObject $dbInfo NoteProperty MountedOnServer $db.MountedOnServer.Split(".")[0]

    if ($db.DatabaseSize -ne $null) {
        $dbInfo.EdbFileSizeInGB = [Math]::Round($db.DatabaseSize.ToBytes()/1GB, 0)
        $totalStorageBytes += $db.DatabaseSize.ToBytes()
    }

    if ($db.LastFullBackup -gt (Get-Date).AddDays(-1)) {
        $dbInfo.BackupStatus = "OK (<24h)"
    } else {
        $dbInfo.BackupStatus = "NOT OK (>24h)"
    }

    $dbInfo.AvailableSpaceInMB = [Math]::Round($db.AvailableNewMailboxSpace.ToBytes()/1MB, 0)
    $dbUsers = $allMailboxes | ? { $_.homeMDB -match "CN=$($db.Name)," }
    $dbInfo.Mailboxes = $dbUsers.Count

    [UInt64]$totalDbUserQuota = 0
    $dbQuota = $db.ProhibitSendReceiveQuota.Value.ToBytes()

    $usersCount = $dbUsers.Count
    $j = 0
    $start = Get-Date
    foreach ($user in $dbUsers) {
        if ($user.mDBUseDefaults -eq $true) {
            $totalDbUserQuota += $dbQuota
        } else {
            $totalDbUserQuota += (Get-Mailbox $user.cn).ProhibitSendReceiveQuota.Value.ToBytes()
        }

        if ($IncludeTopStorageUsers -or $RunSetMailboxQuotaLimits) {
            $stats = Get-MailboxStatistics $user.cn
            if ($stats -ne $null) {
                if ($IncludeTopRecipients) {
                    $userInfoArray[$user.cn] = $stats.TotalItemSize.Value.ToBytes()
                }

                if ($RunSetMailboxQuotaLimits -and
                        (Test-Path "$cwd\Set-MailboxQuotaLimits.ps1") -and
                        ($stats.StorageLimitStatus -eq 'IssueWarning' -or
                         $stats.StorageLimitStatus -eq 'ProhibitSend' -or
                         $stats.StorageLimitStatus -eq 'MailboxDisabled')) {
                    Write-Verbose "Raising quota for user $($user.cn)"
                   & "$cwd\Set-MailboxQuotaLimits.ps1" -Identity $user.cn -Verbose:$Verbose -Confirm:$false
                }
            }
        }
        $j++
        if (($j % 10) -eq 0) {
            $end = Get-Date
            $totalSeconds = ($end - $start).TotalSeconds
            $timePerUser = $totalSeconds / $j
            $timeLeft = $timePerUser * ($usersCount - $j)
            Write-Progress  -Activity "Gathering User Statistics for Database ($db.Name)" `
                            -Status $user.cn `
                            -PercentComplete ($j/$usersCount*100) `
                            -Id 2 -ParentId 1 `
                            -SecondsRemaining $timeLeft
        }
    }
    Write-Progress -Activity "Gathering User Statistics" -Status "Finished" -Id 2 -ParentId 1 -Completed

    $dbInfo.CommitPercent = [Math]::Round(($totalDbUserQuota/$MaxDatabaseSizeInBytes*100), 0)

    Write-Verbose "Database Name: $($dbInfo.Identity)"
    Write-Verbose "Database Size: $($dbInfo.EdbFileSizeInGB)"
    Write-Verbose "Database Available Space: $($dbInfo.AvailableSpaceInMB)MB"
    Write-Verbose "Database Commit %: $($dbInfo.CommitPercent)"
    Write-Verbose "Database Backup Status: $($dbInfo.BackupStatus)"
    Write-Verbose "Database Mounted on:  $($dbInfo.MountedOnServer)"
    $null = $dbInfoArray.Add($dbInfo)
    $i++
}
Write-Progress -Activity "Gathering Database Statistics" -Status "Finished" -Id 1 -Completed

if ($IncludeTopRecipients) {
    Write-Verbose "Getting Top Recipient Counts"
    $recipientCounts = & "$cwd\Get-TopRecipientCounts.ps1"
}

Write-Host "Getting UserMailbox count..."
$statsArray["User Mailboxes"] = ($allMailboxes | ? { $_.msExchRecipientTypeDetails -eq 1 }).Count
Write-Verbose "Found $($statsArray['User Mailboxes']) UserMailboxes"

Write-Host "Getting SharedMailbox count..."
$statsArray["Shared Mailboxes"] = ($allMailboxes | ? { $_.msExchRecipientTypeDetails -eq 4 }).Count
Write-Verbose "Found $($statsArray['Shared Mailboxes']) Shared Mailboxes"

Write-Host "Getting Resource Mailbox count..."
$statsArray["Resource Mailboxes"] = ($allMailboxes | ? { $_.msExchRecipientTypeDetails -eq 16 -or $_.msExchRecipientTypeDetails -eq 32 }).Count
Write-Verbose "Found $($statsArray['Resource Mailboxes']) Resource Mailboxes"

Write-Host "Getting Disabled Mailboxes count..."
$statsArray["Disabled Mailboxes"] = ($allMailboxes | ? { $_.msExchRecipientTypeDetails -eq 1 -and ($_.userAccountControl -band 2) -eq 2}).Count
Write-Verbose "Found $($statsArray['Disabled Mailboxes']) Disabled Mailboxes"

Write-Host "Getting MailUser count..."
$statsArray["Mail Users"] = (adfind -csv -q -b 'OU=JMUma,dc=ad,dc=jmu,dc=edu' -f '(&(!(homeMDB=*))(targetAddress=*))' cn | ConvertFrom-Csv).Count
Write-Verbose "Found $($statsArray['Mail Users']) MailUsers"

Write-Host "Getting Distribution Group count..."
$statsArray["Distribution Groups"] = (Get-DistributionGroup).Count
Write-Verbose "Found $($statsArray['Distribution Groups']) Distribution Groups"

$statsArray["Total Storage Used"] = "{0:N2} GB" -f ($totalStorageBytes/1GB)
Write-Verbose "Total Storage Used:  $($statsArray['Total Storage Used'])"

$Body  = @"
<html>
    <head>
        <style type="text/css">
            td { text-align:right; }
            td.Identity { text-align:left; font-weight:bold; }
            td.server { text-align:center; text-transform:lowercase; }
            td.warning { background-color:yellow; }
            td.error { background-color:red; }
        </style>
    </head>
    <body>

    <h1>Exchange Statistics for $(Get-Date -Format d)</h1>

    <h2>User Statistics</h2>
    <table cellpadding="2">

"@

foreach ($key in $statsArray.Keys) {
    $Body += @"
        <tr>
            <td>$($key):</td>
            <td>$($statsArray[$key])</td>
        </tr>

"@
}

$Body += @"
    </table>
    <br />
    <hr />

"@

if ($IncludeTopRecipients) {
    $Body += @"
    <br/>
    <h2>Top Senders by Total Recipient Count (last 24 hours)</h2>

    <table border="1" cellpadding="2">
        <tr>
            <th>Sender</th>
            <th>Receipients</th>
        </tr>

"@
    for ($i = 0; $i -lt $recipientCounts.Count; $i++) {
        $Body += @"
        <tr>
            <td class="identity">$($recipientCounts[$i].Key)</td>
            <td>$($recipientCounts[$i].Value)</td>
        </tr>

"@
    }
    $Body += @"
    </table>
    <br />
    <hr />

"@
}

if ($IncludeTopStorageUsers) {
    $Body += @"
    <h2>Top Users by Mailbox Size</h2>
    <table border=1 cellpadding=2>
        <tr>
            <th>User</th>
            <th>Mailbox Size</th>
        </tr>

"@

    foreach ($user in ($userInfoArray.GetEnumerator() | Sort Value -Descending | Select -First 10)) {
        $Body += @"
        <tr>
            <td class="identity">$($user.Name)</td>
            <td>$([Math]::Round($user.Value/1MB, 0)) MB</td>
        </tr>
"@
    }

    $Body += @"
    </table>
    <br />
    <hr />

"@

}

$Body += @"
    <br/>
    <h2>Database Information</h2>
    <table border="1" cellpadding="5">
        <tr>
            <th>Database</th>
            <th>Mailbox Count</th>
            <th>EDB File Size</th>
            <th>Available Space</th>
            <th>Commit Percentage</th>
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

    $Body += "`n"

    if ($db.BackupStatus -match 'NOT OK') {
        $Body += '<td class="warning">{0}</td>' -f $($db.BackupStatus)
    } else {
        $Body += '<td>{0}</td>' -f $($db.BackupStatus)
    }

    $Body += @"

            <td class="server">$($db.MountedOnServer)</td>
        </tr>

"@
}

$Body += @"
    </table>
    <br />
    <hr />
    <br />
    </body>
</html>
"@

Send-MailMessage -From $From -To $To -Subject $Title -Body $Body -SmtpServer $SmtpServer -UseSsl -BodyAsHtml
Write-Verbose "Sent Email"
