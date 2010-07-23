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

# Change these to suit your environment
$SmtpServer = "it-exhub.ad.jmu.edu"
$From       = "it-exmaint@jmu.edu"
$To         = "wrightst@jmu.edu, gumgs@jmu.edu, liskeygn@jmu.edu, stockntl@jmu.edu, flynngn@jmu.edu, najdziav@jmu.edu"
#$To         = "wrightst@jmu.edu"
$Title      = "Exchange User Detail for $(Get-Date -Format d)"

##################################
$cwd = [System.IO.Path]::GetDirectoryName(($MyInvocation.MyCommand).Definition)

$totalStorage = 0
$dbInfoArray = New-Object System.Collections.ArrayList

foreach ($db in (Get-MailboxDatabase -IncludePreExchange2010 -Status)) { 
    Write-Host "Processing $db"
    $dbInfo = New-Object PSObject
    $dbInfo = Add-Member -PassThru -InputObject $dbInfo NoteProperty Identity $db.Name
    Add-Member -InputObject $dbInfo NoteProperty Size $null
    Add-Member -InputObject $dbInfo NoteProperty BackupStatus $null
    Add-Member -InputObject $dbInfo NoteProperty LastFullBackup $db.LastFullBackup

    if ($db.DatabaseSize -ne $null) {
        $dbInfo.Size = $db.DatabaseSize.ToMB()
        $totalStorage += $db.DatabaseSize.ToMB()
    }

    if ($db.LastFullBackup -gt (Get-Date).AddDays(-1)) {
        $dbInfo.BackupStatus = "OK (<24h)"
    } else {
        $dbInfo.BackupStatus = "NOT OK (>24h)"
    }

    $null = $dbInfoArray.Add($dbInfo)
}

$recipientCounts = & "$cwd\Get-TopRecipientCounts.ps1"

Write-Host "Getting Distribution Group count..."
$dls = (Get-DistributionGroup).Count
Write-Host "Getting MailUser count..."
$mailusers = $(adfind -q -b 'OU=JMUma,dc=ad,dc=jmu,dc=edu' -c -f targetAddress=*)[1].Split(" ")[0]
Write-Host "Getting UserMailbox count..."
$users = $(adfind -q -b 'OU=JMUma,dc=ad,dc=jmu,dc=edu' -c -f homeMDB=*)[1].Split(" ")[0]
Write-Host "Getting SharedMailbox count..."
$shared = (Get-User -ResultSize Unlimited -RecipientTypeDetails SharedMailbox).Count
Write-Host "Getting Resource Mailbox count..."
$resources = (Get-User -ResultSize Unlimited -RecipientTypeDetails RoomMailbox,EquipmentMailbox).Count

$Body  = @"
     User Mailboxes:	$users
   Shared Mailboxes:	$shared
 Resource Mailboxes:	$resources
         Mail Users:	$mailusers
Distribution Groups:	$dls
Total Storage Used*:	$totalStorage MB
(*only for E2010 databases)
-----------------------------------------------------------------------

Top Senders by Total Recipient Count (last 24 hours):
$($recipientCounts | ft -Autosize | Out-String)
-----------------------------------------------------------------------

Database Information:
$($dbInfoArray | Sort Identity | ft -AutoSize | Out-String)
-----------------------------------------------------------------------
"@

& "$cwd\Send-Email.ps1" -From $From -To $To -Subject $Title -Body $Body -SmtpServer $SmtpServer 

