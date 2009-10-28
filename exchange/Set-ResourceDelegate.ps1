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

param ([string]$Identity, [string]$Owner, [switch]$EmailOwner=$true)

if ($Owner -eq '' -or $Identity -eq '') {
    Write-Host "Please specify the Identity and Owner"
    return
}

$DomainController = "jmuadc1.ad.jmu.edu"
$resource = Get-Mailbox $Identity
$delegate = Get-Mailbox $Owner

if ($resource -eq $null) {
    Write-Error "Could not find Resource"
    return
}

Write-Host "Setting Calendar Settings: "

# Grant Send-As rights to the owner:
$resource | Add-ADPermission -ExtendedRights "Send-As" -User $owner `
            -DomainController $DomainController

# Give the owner Full Access to the resource:
$resource | Add-MailboxPermission -DomainController $DomainController `
            -AccessRights FullAccess -User $Owner

# Grant SendOnBehalfOf rights to the owner:
$sobo = (Get-Mailbox -DomainController $DomainController -Identity $resource).GrantSendOnBehalfTo
if ( !$sobo.Contains((Get-User $Owner).DistinguishedName) ) {
    $sobo.Add( (Get-User $Owner).DistinguishedName )
}
$resource | Set-Mailbox -DomainController $DomainController `
            -GrantSendOnBehalfTo $sobo

# Set the ResourceDelegates
$resourceDelegates = (Get-MailboxCalendarSettings -Identity $resource).ResourceDelegates
if ( !($resourceDelegates.Contains((Get-User $Owner).DistinguishedName)) ) {
    $resourceDelegates.Add( (Get-User $Owner).DistinguishedName )
}

foreach ($i in 1..10) {
    $error.Clear()
    $resource | Set-MailboxCalendarSettings -DomainController $DomainController `
                -AllRequestOutOfPolicy:$True -AutomateProcessing AutoAccept `
                -BookingWindowInDays 365 -ResourceDelegates $resourceDelegates `
                -ErrorAction SilentlyContinue
    if (![String]::IsNullOrEmpty($error[0])) {
        Write-Host -NoNewLine "."
        Start-Sleep $i
    } else {
        Write-Host "done."
        break
    }
}

if ($EmailOwner) {
    $From   = "Seth Wright <wrightst@jmu.edu>"
    $Cc     = "wrightst@jmu.edu, boyledj@jmu.edu, millerca@jmu.edu"
    $Title = "Information about Exchange resource `"$resource`""
    $To = $delegate.PrimarySmtpAddress.ToString()

    $Body = @"
You have been identified as a resource owner / delegate for the
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

    https://exchange.jmu.edu/owa/$($resource.PrimarySMTPAddress)`n

(Log in using your own eID and password.)`n

Click on the Options link in the upper-right-hand corner, then click the
"Resource Settings" option in the left-hand column.  Most of the options
should be self-explanatory.  For instance, if you would like to alter
the settings of this resource such that no user can automatically book
it, and that every request must be approved, simply change both settings
that start with "These users can schedule automatically..." to "Select
users and groups" instead of "Everyone", and set "These users can submit
a request for manual approval..." to "Everyone".`n

If you have any questions, please contact the JMU Computing HelpDesk at
helpdesk@jmu.edu, or by phone at 540-568-3555.
"@

    $SmtpClient = New-Object System.Net.Mail.SmtpClient
    $SmtpClient.Host = "it-exhub.ad.jmu.edu"
    $SmtpClient.Port = 25
    $Message = New-Object System.Net.Mail.MailMessage $From, $To, $Title, $Body
    $Message.Cc.Add($Cc)
    $SmtpClient.Credentials = [System.Net.CredentialCache]::DefaultNetworkCredentials
    $SmtpClient.Send($message)
    Write-Output "Sent message to $To for resource `"$resource`""
}
