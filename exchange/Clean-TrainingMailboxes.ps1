################################################################################
# 
# $URL$
# $Author$
# $Date$
# $Rev$
# 
# DESCRIPTION:  Creates mailboxes in Exchange
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

# Change these to suit your environment
$SmtpServer = "it-exhub.ad.jmu.edu"
$From       = "it-exmaint@jmu.edu"
$To         = "wrightst@jmu.edu, millerca@jmu.edu"
$Title      = "Training Mailboxes"

##################################
$cwd = [System.IO.Path]::GetDirectoryName(($MyInvocation.MyCommand).Definition)
$now = Get-Date

# Get all of the training user mailboxes
$mboxes = Get-Mailbox -Filter { Name -like "Training User*" } -RecipientTypeDetails UserMailbox
# Disable them (remove the mailbox, but not the user account)
$mboxes | Disable-Mailbox -Confirm:$false
# Sleep for a while.
Start-Sleep 120
# Recreate the mailboxes
$Body = $mboxes | Enable-Mailbox -Database it-exmbx1\Training -ManagedFolderMailboxPolicy "Default Managed Folder Policy" -ManagedFolderMailboxPolicyAllowed:$true | Out-String

& "$cwd\Send-Email.ps1" -From $From -To $To -Subject $Title -Body $Body -SmtpServer $SmtpServer 

