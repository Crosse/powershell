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

$Body  = @"
Total Mailbox Count:	$($everything.Count)
     User Mailboxes:	$($users.Count)
   Shared Mailboxes:	$shared
          Resources:	$resources
Distribution Groups:	$dls

 Training Mailboxes:	$ittraining
 HelpDesk Mailboxes:	$helpdesk

Known Phishing Addresses Blacklisted:  $antiphishing
"@

$SmtpServer = "it-exhub.ad.jmu.edu"
$SmtpClient = New-Object System.Net.Mail.SmtpClient
$SmtpClient.Credentials = [System.Net.CredentialCache]::DefaultNetworkCredentials
$SmtpClient.Port = 25
$SmtpClient.host = $SmtpServer
$SmtpClient.Send($From, $To, $Title, $Body)

#stop-transcript

