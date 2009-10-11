################################################################################
# 
# $URL$
# $Author$
# $Date$
# $Rev$
# 
# DESCRIPTION:  Creates a new resource in accordance with JMU's current naming
#               policies, etc.
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

param ( [string]$DisplayName, 
        [string]$Owner, 
        [switch]$Room,
        [switch]$Equipment )

# Change these to suit your environment
$SmtpServer = "it-exhub.ad.jmu.edu"
$From       = "Seth Wright <wrightst@jmu.edu>"
$Cc         = "wrightst@jmu.edu" #, gumgs@jmu.edu, boyledj@jmu.edu"
$Fqdn       = "exchange.jmu.edu"
$DomainController = "jmuadc4.ad.jmu.edu"
# TODO: When automatic SG determination is done, rewrite this line to 
# use that script.
$Database   = "IT-ExMbx1\Pilot"
$BaseDN     = "ad.jmu.edu/ExchangeObjects/Resources"

##################################

$SmtpClient = New-Object System.Net.Mail.SmtpClient
$SmtpClient.Host = $SmtpServer
$SmtpClient.Port = 25
$SmtpClient.Credentials = [System.Net.CredentialCache]::DefaultNetworkCredentials

if ( $DisplayName -eq '' -or $Owner -eq '') {
Write-Output "-DisplayName and -Owner are required"
    return
}

if ( !($Room -or $Equipment) ) {
    Write-Output "Please specify either -Room or -Equipment"
    return
}

$ou = $BaseDN
if  ( $Room ) {
    $ou += "/Rooms"
} elseif ( $Equipment ) {
    $ou += "/Equipment"
}

$Name  = $DisplayName
$alias = $DisplayName
$alias = $alias.Replace('Conference Room', 'ConfRoom')
$alias = $alias.Replace('Lecture Hall', 'LectureHall')
$alias = $alias.Replace(' Hall', '')
$alias = $alias.Replace(' ', '_')

$cmd  = "New-Mailbox -DomainController $DomainController -Database `"$Database`""
$cmd += "-OrganizationalUnit `"$ou`" -Name `"$Name`" -Alias `"$alias`" -UserPrincipalName "
$cmd += "`"$($alias)@ad.jmu.edu`" -DisplayName `"$DisplayName`""

$error.Clear()

if ( $Room ) {
    $cmd += " -Room"
} elseif ( $Equipment ) {
    $cmd += " -Equipment"
}

Invoke-Expression($cmd)

if (!([String]::IsNullOrEmpty($error[0]))) {
    return
}

$resource = Get-Mailbox -DomainController $DomainController -Identity $alias

if ( !$resource) {
    Write-Output "Could not find $alias in Active Directory."
    return
}

# Give the owner Full Access to the resource:
Add-MailboxPermission -DomainController $DomainController `
    -Identity $alias -AccessRights FullAccess -User $Owner

# Grant SendOnBehalfOf rights to the owner:
Set-Mailbox -DomainController $DomainController -Identity $alias `
    -GrantSendOnBehalfTo $owner

# Send a message to the mailbox.  Somehow this helps...but sleep first.
Write-Host "Blocking for 60 seconds for the mailbox to be created:"
foreach ($i in 1..60) { 
    Write-Host -NoNewLine "."
    Start-Sleep 1
}

Write-Host "done.`nSending a message to the resource to initialize the mailbox."
$Message = New-Object System.Net.Mail.MailMessage "wrightst@jmu.edu", "$($alias)@ad.jmu.edu", "disregard", "disregard"
$SmtpClient.Send($message)

# Set the default calendar settings on the resource:
# Unfortunately, this fails if the mailbox isn't fully created yet, so introduce a wait.
Write-Host "Setting Calendar Settings: "

foreach ($i in 1..10) {
    $error.Clear()
    Set-MailboxCalendarSettings -DomainController $DomainController `
        -Identity $alias -AllRequestOutOfPolicy:$True -AutomateProcessing AutoAccept `
        -BookingWindowInDays 365 -ResourceDelegates $owner -ErrorAction SilentlyContinue
    if (![String]::IsNullOrEmpty($error[0])) {
        Write-Host -NoNewLine "."
        Start-Sleep $i
    } else {
        Write-Host "done."
        break
    }
}

if ( !(Get-MailboxCalendarSettings -DomainController $DomainController -Identity $resource).ResourceDelegates ) {
    Write-Output "Skipping `"$resource`" because it has no delegates"
    continue
}
$Title = "Information about Exchange resource `"$resource`""
$To = "$($owner)@jmu.edu"

$Body = @"
You have been identified as the resource owner / delegate for the
following Exchange resource:`n

    $resource`n

This email is to inform you about the booking policy for this resource,
and how you can change it if the defaults do not suit the resource.
Currently, the resource will automatically accept booking requests
that do not conflict with other bookings, and will require your approval
if a request is made that conflicts with another booking.`n

If you would like to change this behavior, you may do so by using
Outlook Web Access.  Open Internet Explorer and navigate to the
following URL:`n

    https://$($Fqdn)/owa/$($resource.PrimarySMTPAddress)`n

(Log in using your own eID and password.)`n

Click on the Options link in the upper-right-hand corner, then click the
"Resource Settings" option in the left-hand column.  Most of the options
should be self-explanatory.  For instance, if you would like to alter
the settings of this resource such that no user can automatically book
it, and that every request must be approved, simply change both settings
that start with "These users can schedule automatically..." to "Select
users and groups" instead of "Everyone", and set "These users can submit
a request for manual approval..." to "Everyone".`n

If you have any questions, please let me know.
"@

$Message = New-Object System.Net.Mail.MailMessage $From, $To, $Title, $Body
$Message.Cc.Add($Cc)
$SmtpClient.Send($message)
Write-Output "Sent message to $To for resource `"$resource`""

