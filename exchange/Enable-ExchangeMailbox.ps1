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

param ( $User="",
        $Server="localhost",
        [switch]$Automated=$false,
        [switch]$Force=$false,
        [switch]$Verbose=$false,
        [System.Collections.Hashtable]$Databases=$null,
        $inputObject=$null )

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

    $cwd = [System.IO.Path]::GetDirectoryName(($MyInvocation.MyCommand).Definition)

    if ($Databases -eq $null) {
        $dbs = & "$cwd\Get-BestDatabase.ps1" -Server $Server -Single:$false
        if ($dbs -eq $null) {
            Write-Error "Could not enumerate databases!"
            return
        }
    } else {
        $dbs = $Databases
    }

    $exitCode = 0
} # end 'BEGIN{}'

# This section executes for each object in the pipeline.
PROCESS {
    if ( !($_) -and !($User) ) { 
        Write-Output "No user given."
        return
    }

    if ($_) { $User = $_ }

# Was a username passed to us?  If not, bail.
    if ([String]::IsNullOrEmpty($User)) { 
        Write-Error "USAGE:  Enable-ExchangeMailbox -User `$User"
        return
    }

    $objUser = Get-User $User -ErrorAction SilentlyContinue

    if (!($objUser)) {
        Write-Output "$User`tis not a valid user in Active Directory."
        $exitCode += 1
        return
    } else { 
        switch ($objUser.RecipientTypeDetails) {
            'User' { break }
            'MailUser' { break }
            'UserMailbox' {
                if (!$Automated) {
                    Write-Output "$($objUser.SamAccountName)`talready has a mailbox"
                }
                return
            }
            default {
                Write-Output "$($objUser.SamAccountName)`tis a $($objUser.RecipientTypeDetails) object, refusing to create mailbox"
                $exitCode += 1
                return
            }
        }
    }

# Don't auto-create mailboxes for users in the Students OU
    if ($objUser.DistinguishedName -match 'Student') {
        Write-Output "$($User)`tis listed as a student, refusing to create mailbox"
        $exitCode += 1
        return
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
        & "$cwd\Deprovision-User.ps1" -User $objUser.DistinguishedName -Confirm:$false
        if ($LASTEXITCODE -gt 0) {
            Write-Output "An error occurred; refusing to create mailbox."
            $exitCode += 1
            return
        }
    }

# Enable the mailbox
    $Error.Clear()
    Enable-Mailbox -Database "$($candidate)" -Identity $objUser `
        -ManagedFolderMailboxPolicy "Default Managed Folder Policy" `
        -ManagedFolderMailboxPolicyAllowed:$true `
        -DomainController $DomainController -ErrorAction SilentlyContinue

    if ($Error[0] -ne $null) {
        $exitCode += 1
        Write-Output $Error[0]
        return
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
    exit $exitCode
} # end 'END{}'

