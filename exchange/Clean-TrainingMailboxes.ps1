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
