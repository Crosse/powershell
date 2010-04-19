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

param ( $User="", [string]$Server="localhost", [switch]$Verbose=$false, $inputObject=$null )

# This section executes only once, before the pipeline.
BEGIN {
    if ($inputObject) {
        Write-Output $inputObject | &($MyInvocation.InvocationName)
        break
    }

    $srv = Get-ExchangeServer $Server -ErrorAction SilentlyContinue
    if ($srv -eq $null) {
        Write-Error "Could not find Exchange Server $Server"
        exit
    }

    $cwd                = [System.IO.Path]::GetDirectoryName(($MyInvocation.MyCommand).Definition)
    $Users              = New-Object System.Collections.ArrayList
    $DomainController   = (gc Env:\LOGONSERVER).Replace('\', '')

    if ($DomainController -eq $null) { 
        Write-Warning "Could not determine the local computer's logon server!"
        return
    }

}

PROCESS {
    if ( !($_) -and !($User) ) { 
        Write-Error "No user given."
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
        return
    } else { 
        if ($objUser.RecipientTypeDetails -ne 'User' -and 
                $objUser.RecipientTypeDetails -ne 'MailUser') {
            # If the user is already a Mailbox, don't warn on that.
            if ($objUser.RecipientTypeDetails -notmatch 'Mailbox') {
                Write-Output "$($objUser.SamAccountName)`tis a $($objUser.RecipientTypeDetails) object, refusing to create mailbox"
            }
            return
        }
    }

# Don't auto-create mailboxes for users in the Students OU
    if ($objUser.DistinguishedName -match 'Student') {
        Write-Output "$($User)`tis listed as a student, refusing to create mailbox"
        return
    }

# Add the user to the queue.
    $null = $Users.Add($objUser)
} # end 'PROCESS{}'

# This section executes only once, after the pipeline.
END {
    $now = Get-Date
    if ($Users.Count -gt 0) {
        $Users | & "$cwd\Enable-ExchangeMailbox.ps1" -Server $Server -Verbose:$Verbose | 
            Out-String |
            & "$cwd\Send-Email.ps1" -From 'it-exmaint@jmu.edu' -To 'wrightst@jmu.edu' `
                -Subject "Exchange Provisioning $($now.ToString()) ($($Users.Count) Users)" `
                -SmtpServer it-exhub.ad.jmu.edu
    } else {
        & "$cwd\Send-Email.ps1" -From 'it-exmaint@jmu.edu' -To 'wrightst@jmu.edu' `
            -Subject "Exchange Provisioning $($now.ToString()) (Nothing to do)" `
            -Body 'No users to provision.' -SmtpServer it-exhub.ad.jmu.edu
    }
} # end 'END{}'
