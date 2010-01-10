################################################################################
# 
# $URL$
# $Author$
# $Date$
# $Rev$
# 
# DESCRIPTION:  Returns the best database in which to create a new mailbox.
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

param ([string]$Server)

if ($Server -eq '') {
    Write-Host "Please specify the Server"
    return
}

##################################

$srv = Get-ExchangeServer $Server
if ($srv -eq $null) {
    Write-Error "Could not find Exchange Server $Server"
    return
}

$databases = Get-MailboxDatabase -Server $srv | Where { $_.Name -notmatch 'Training' }
if ($databases -eq $null) {
    Write-Error "Could not enumerate databases on server $Server"
    return
}

$candidateUserCount = -1
$candidate = ""

foreach ($database in $databases) {
    $mailboxCount = (Get-Mailbox -Database $database).Count
    if ($? -eq $False) {
        Write-Error "Error processing database $database"
        return
    }

    if ($mailboxCount -lt $candidateUserCount -or $candidateUserCount -eq -1) {
        Write-Host -NoNewLine "!"
        $candidateUserCount = $mailboxCount
        $candidate = $database
    } else {
        Write-Host -NoNewLine "."
    }

}

Write-Host "`nCandidate Database: $candidate contains $candidateUserCount mailboxes"
$candidate
