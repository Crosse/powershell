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

param ( $User="", $Server="localhost", [switch]$Verbose=$false, $inputObject=$null )

# This section executes only once, before the pipeline.
BEGIN {
    if ($inputObject) {
        Write-Output $inputObject | &($MyInvocation.InvocationName)
        break
    }

    $srv = Get-ExchangeServer $Server
    if ($srv -eq $null) {
        Write-Error "Could not find Exchange Server $Server"
        return
    }

    $DomainController = (gc Env:\LOGONSERVER).Replace('\', '')
    if ($DomainController -eq $null) { 
        Write-Warning "Could not determine the local computer's logon server!"
        return
    }

    $databases = Get-MailboxDatabase -Server $srv -Status | 
    Where { $_.Mounted -eq $True -and $_.Name -match "^SG" }
    if ($databases -eq $null) {
        Write-Error "Could not enumerate databases on server $Server"
        return
    }

    $i = 1
    $dbs = New-Object System.Collections.Hashtable
    Write-Host -NoNewLine "["
    foreach ($database in $databases) {
        if ($i % 10 -eq 0) {
            Write-Host -NoNewLine "|"
        } elseif ($i % 5 -eq 0) {
            Write-Host -NoNewLine "+"
        } else {
            Write-Host -NoNewLine "."
        }
        $i++

        $mailboxCount = (Get-Mailbox -Database $database).Count
        if ($? -eq $False) {
            Write-Error "Error processing database $database"
            return
        }

        $maxUsers = 200GB / (Get-MailboxDatabase $database).ProhibitSendReceiveQuota.Value.ToBytes()

        if ($mailboxCount -le $maxUsers) {
            # Normally we'd not add this database
            # if the mailboxCount was greater than the maximum
            # number of users allowed for the database,
            # but we're fudging it for a while.
        }
        $dbs.Add($database.Identity.ToString(), $mailboxCount)
    }
    Write-Host "]"

} # end 'BEGIN{}'

# This section executes for each object in the pipeline.
PROCESS {
    if ( !($_) -and !($User) ) { 
        Write-Output "No user given."
        return
    }

    if ($_) { $User = $_ }

# Was a username passed to us?  If not, bail.
    if (!($User)) { 
        Write-Output "USAGE:  Enable-ExchangeMailbox -User `$User"
    }

    $objUser = Get-User $User -ErrorAction SilentlyContinue

    if (!($objUser)) {
        Write-Output "$User is not a valid user in Active Directory."
        return
    } else { 
        if ($objUser.RecipientTypeDetails -ne 'User' -and 
                $objUser.RecipientTypeDetails -ne 'MailUser') {
            Write-Output "$($objUser): cannot operate on $($objUser.RecipientTypeDetails) objects"
            return
        }
    }

    $candidate = $null
    foreach ($db in $dbs.Keys) {
        if ($candidate -eq $null) {
            $candidate = $db
        } else {
            if ($dbs[$db] -lt $dbs[$candidate]) {
                $candidate = $db
            }
        }
    }

    if ($Verbose) {
        Write-Output "Assigning $($objUser.SamAccountName) to database $candidate"
    }

# Save this off because Exchange blanks it out...
    $displayNamePrintable = $objUser.SimpleDisplayName

# If the user is a MailUser already, remove the Exchange bits first
    if ($objUser.RecipientTypeDetails -match 'MailUser') {
        if ($Verbose) {
            Write-Output "User is a MailUser; running Disable-MailUser first"
        }
        $error.Clear()
        Disable-MailUser -Identity $objUser.DistinguishedName -Confirm:$false `
            -DomainController $DomainController -ErrorAction SilentlyContinue
        if (![String]::IsNullOrEmpty($error[0])) {
            Write-Output $error[0]
        }
    }

# Enable the mailbox
    $Error.Clear()
    Enable-Mailbox -Database "$($candidate)" -Identity $objUser `
    -ManagedFolderMailboxPolicy "Default Managed Folder Policy" `
    -ManagedFolderMailboxPolicyAllowed:$true `
    -DomainController $DomainController -ErrorAction SilentlyContinue
    if ($Error[0] -ne $null) {
        Write-Output $Error[0]
    } else {
# No error, so set the SimpleDisplayName now that Exchange has 
# helpfully removed it.
        if ($Verbose) {
            Write-Output "Resetting $($objUser.SamAccountName)'s SimpleDisplayName to `"$displayNamePrintable`""
        }
        $error.Clear()
        Set-User $objUser -SimpleDisplayName "$($displayNamePrintable)" `
            -DomainController $DomainController -ErrorAction SilentlyContinue
        if (![String]::IsNullOrEmpty($error[0])) {
            Write-Output $error[0]
        }

# Increment the running mailbox total for the candidate database.
        $dbs[$candidate]++
    }

} # end 'PROCESS{}'

# This section executes only once, after the pipeline.
END {
} # end 'END{}'

