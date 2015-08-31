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
$SmtpServer = "mailgw.jmu.edu"
$From       = "it-exmaint@jmu.edu"
$To         = "wrightst@jmu.edu, millerca@jmu.edu"
$Title      = "Training Mailboxes"

$DomainController = "jmuadc1.ad.jmu.edu"

##################################

$cwd = [System.IO.Path]::GetDirectoryName(($MyInvocation.MyCommand).Definition)

$now = Get-Date
Start-Transcript "cleanTrainingMailboxes_$($now.Year)-$($now.Month)-$($now.Day).log" -Append

$results = ""

# Get all of the training user mailboxes
$trainingUsers = Get-User -OrganizationalUnit ad.jmu.edu/ExchangeObjects/ITTraining -Filter { Name -like "Training User*" } | Sort UserPrincipalName

foreach ($user in $trainingUsers) {
  $results += "Processing $($user.Name)..."
  if ($user.RecipientTypeDetails -match 'UserMailbox') {
    $oldDatabase = (Get-Mailbox $user.Name).Database
    $error.clear()
    Disable-Mailbox -Identity $user.Name -Confirm:$false -DomainController $DomainController
    if (![String]::IsNullOrEmpty($error[0])) {
      $results += "`n==> An error occurred disabling the mailbox:  $($error[0])"
      continue
    }
  }

  $error.clear()
  Enable-Mailbox -Identity $user `
    -ManagedFolderMailboxPolicy "Default Managed Folder Policy" `
    -ManagedFolderMailboxPolicyAllowed:$true `
    -Alias $user.SamAccountName `
    -DomainController $DomainController
  if (![String]::IsNullOrEmpty($error[0])) {
    $results += "`n==> An error occurred enabling the mailbox:  $($error[0])"
  } else {
    $results += "done.`n"
  }
}

& "$cwd\Send-Email.ps1" -From $From -To $To -Subject $Title -Body $results -SmtpServer $SmtpServer 

Stop-Transcript