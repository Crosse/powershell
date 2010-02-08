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

$now = Get-Date
$logFile = "training_$($now.Year)-$($now.Month)-$($now.Day).log"
Start-Transcript $logFile

# Get all of the training user mailboxes
$mboxes = Get-Mailbox -Filter { Name -like "Training User*" } -RecipientTypeDetails UserMailbox
# Disable them (remove the mailbox, but not the user account)
$mboxes | Disable-Mailbox -Confirm:$false
# Sleep for a while.
Start-Sleep 120
# Recreate the mailboxes
$mboxes | Enable-Mailbox -Database it-exmbx1\Training -ManagedFolderMailboxPolicy "Default Managed Folder Policy" -ManagedFolderMailboxPolicyAllowed:$true
Stop-Transcript

# Send the log as an email attachment
E:\Scripts\Send-Email.ps1 -From '"Exchange Maintenance Account" <it-exmaint@jmu.edu>' -To 'millerca@jmu.edu,moellejf@jmu.edu' -Cc 'wrightst@jmu.edu,najdziav@jmu.edu' -Subject 'Output of Training Mailbox Creation' -Body 'See attachment' -SmtpServer et.jmu.edu -SmtpPort 25 -AttachmentFile $logFile
