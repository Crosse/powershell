################################################################################
# 
# $Id$
# 
# DESCRIPTION:  Sends an email.
#
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

param ( [string]$From='', 
        [string]$To='', 
        [string]$Cc='',
        [string]$Bcc='',
        [string]$Subject,
        [string]$Body='',
        [string]$SmtpServer='',
        [int]$SmtpPort=25,
        [string]$AttachmentFile,
        $Credential=[System.Net.CredentialCache]::DefaultNetworkCredentials)

if ([System.String]::IsNullOrEmpty($From)) {
    Write-Error "Please provide the From: value"
    return
}

if ([System.String]::IsNullOrEmpty($To) -and 
        ([System.String]::IsNullOrEmpty($Cc) -and 
         [System.String]::IsNullOrEmpty($Bcc))) {
    Write-Error "Please provide the To: value"
    return
}
if ([System.String]::IsNullOrEmpty($SmtpServer)) {
    Write-Error "Please provide an SMTP server"
    return
}

if ($Body -eq '') {
    Write-Warning "Sending email with null body!"
}
if ($Subject -eq '') {
    Write-Warning "Sending email with null subject line!"
}

$SmtpClient = New-Object System.Net.Mail.SmtpClient
$SmtpClient.Host = $SmtpServer
$SmtpClient.Port = $SmtpPort
$SmtpClient.Credentials = [System.Net.CredentialCache]::DefaultNetworkCredentials
$Message = New-Object System.Net.Mail.MailMessage 

$Message.From       = $From
$Message.Subject    = $Subject
$Message.Body       = $Body

if (![System.String]::IsNullOrEmpty($To)) {
    $Message.To.Add($To)
}

if (($AttachmentFile -ne $null -and $AttachmentFile -ne '') -and (Test-Path "$AttachmentFile")) {
    $Attachment = New-Object Net.Mail.Attachment((Get-Item $AttachmentFile).ToString())
    $Message.Attachments.Add($Attachment)
}

if (![System.String]::IsNullOrEmpty($Cc)) {
    $Message.Cc.Add($Cc)
}

if (![System.String]::IsNullOrEmpty($Bcc)) {
    $Message.Bcc.Add($Bcc)
}

Write-Output "Sending email from $From with subject line `"$Subject`""
$error.Clear()
$SmtpClient.Send($message)
if ([System.String]::IsNullOrEmpty($error[0])) {
    Write-Output "Message successfully sent."
} else {
    Write-Output "Message send failure."
}

if ($Attachment -ne $null) {
    $Attachment.Dispose()
}

