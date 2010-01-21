################################################################################
# 
# $Id$
# 
# DESCRIPTION:  Sends an email with relevant Exchange statistics to various
#               users.
#
# Copyright (c) 2009 Seth Wright (wrightst@jmu.edu)
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
################################################################################

#start-transcript exchangeinfo.log

$From = "Exchange System <it-exmaint@jmu.edu>"
$To = "wrightst@jmu.edu, gumgs@jmu.edu, liskeygn@jmu.edu, stockntl@jmu.edu, boyledj@jmu.edu, millerca@jmu.edu, flynngn@jmu.edu, kingms@jmu.edu, mastrw@jmu.edu, hardbaht@jmu.edu"
$Title = "Exchange User Detail for $(Get-Date -Format d)"

$everything   = Get-Mailbox -Server it-exmbx1 -ResultSize Unlimited
$allMailboxes = $everything | where { $_.OrganizationalUnit -notmatch 'ITTraining' -and $_.OrganizationalUnit -notmatch 'HelpDesk' }

$ittraining   = ($everything | where { $_.OrganizationalUnit -match 'ITTraining' }).Count
$helpdesk     = ($everything | where { $_.OrganizationalUnit -match 'HelpDesk' }).Count

$resources    = ($allMailboxes | Where { $_.RecipientTypeDetails -match "RoomMailbox" -or $_.RecipientTypeDetails -match "EquipmentMailbox" }).Count
$shared       = ($allMailboxes | Where { $_.RecipientTypeDetails -match "SharedMailbox" }).Count

$users        = $allMailboxes | where { $_.RecipientTypeDetails -match 'UserMailbox' -and $_.OrganizationalUnit -match "JMUma" } | Sort

$ou           = [ADSI]"LDAP://OU=PhishingReplyAddresses,OU=ExchangeObjects,DC=ad,DC=jmu,DC=edu"
$antiphishing = 0
foreach ($entry in $ou.psbase.Children) { $antiphishing++ }
Remove-Variable ou

$dls = (Get-DistributionGroup).Count

$dbs = New-Object System.Text.StringBuilder
$everything | Group-Object Database | Sort Count,Name -Descending | % { $null = $dbs.AppendFormat("{0,9}`t{1}`n", $_.Count, $_.Name) }

$Body  = @"
Total Mailbox Count:	$($everything.Count)
     User Mailboxes:	$($users.Count)
   Shared Mailboxes:	$shared
          Resources:	$resources
Distribution Groups:	$dls

 Training Mailboxes:	$ittraining
 HelpDesk Mailboxes:	$helpdesk

Known Phishing Addresses Blacklisted:  $antiphishing


Database Utilization:

Mailboxes`tDatabase
---------`t--------
$($dbs.ToString())
"@

$SmtpServer = "it-exhub.ad.jmu.edu"
$SmtpClient = New-Object System.Net.Mail.SmtpClient
$SmtpClient.Credentials = [System.Net.CredentialCache]::DefaultNetworkCredentials
$SmtpClient.Port = 25
$SmtpClient.host = $SmtpServer
$SmtpClient.Send($From, $To, $Title, $Body)

#stop-transcript

