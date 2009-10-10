################################################################################
# 
# $URL$
# $Author$
# $Date$
# $Rev$
# 
# DESCRIPTION: Creates mailboxes in Exchange Labs / Outlook Live
# 
# Copyright (c) 2009 Seth Wright
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
# Revision History:
# 2009-04-06:  - Updated for use with Outlook Live, which needs
#                 PowerShell 2.0 CTP3 and WinRM 2.0 CTP3
#              - Removed dependency on PSCX extensions in favor of
#                 Quest's extensions, since they are 64-bit aware.
#
################################################################################

if (Test-Path function:New-OutlookLiveMailbox) { Remove-Item function:New-OutlookLiveMailbox }

function global:New-OutlookLiveMailbox( $User="", $Credential="", [switch]$f, $inputObject=$Null ) {
    # This section executes only once, before the pipeline.
    BEGIN {
        if ($inputObject) {
            Write-Output $inputObject | &($MyInvocation.InvocationName) -Credential $Credential
            break
        }
    # Check to ensure that the Quest snapin has been registered.
    # Iterate through all the loaded snapins, searching for the Quest snapin.
    foreach ($snapin in (Get-PSSnapin | Sort-Object -Property Name)) {
        if ($snapin.name.ToUpper() -eq "QUEST.ACTIVEROLES.ADMANAGEMENT") {
        # Done, we have the extension and it's loaded.
        $questLoaded = $True
        break
        }
    }
    if (!($questLoaded)) {
        # The Quest snapin was not loaded, so see if the 
        # extension is at least registered with the system.
        foreach ($snapin in (Get-PSSnapin -registered | Sort-Object -Property Name)) {
            if ($snapin.name.ToUpper() -eq "QUEST.ACTIVEROLES.ADMANAGEMENT") {
                # Found the snapin; add it to the environment.
                trap { continue }
                Add-PSSnapin Quest.ActiveRoles.ADManagement
                Write-Host "Quest Active Directory Management Extensions found and added to this session."
                $questLoaded = $True
                break
            }
        }
    }
    
    if (!($questLoaded)) {
        # The Quest snapin is not installed on this system.
        # Print an error and bail.
        Write-Error -Category NotInstalled `
        -RecommendedAction "Install Quest Active Directory Management Extensions" `
        -Message "Quest Active Directory Management Extensions are not installed.  Please install the Extensions and re-run this command."
        continue
    }
        
    # If no credentials were given, ask for them.
    if (!($Credential)) { 
        # This next statement gets rid of the ugly error message that ensues
        # if you cancel the dialog box.  We'll handle that condition next.
        trap { continue }
        # Ask the user for credentials.
        $Credential = Get-Credential
     
    }
    
    if (!($Credential)) {
        # The user cancelled the Get-Credential dialog box.
        # We should probably die now.
        Write-Error -Message "No Credentials were given."
        continue
    }
    
    # This section sets up the connection to Exchange Labs.
    # Only want to do this once for all objects in the pipeline,
    # otherwise this script would take forever.
    # The runspace is destroyed in the END section.
    Write-Host -NoNewline "Attempting to create a runspace for Exchange Labs..."
    $error.Clear()
    $script:ExchangeLabsRS = New-PSSession -ConfigurationName Microsoft.Exchange `
        -ConnectionUri https://ps.outlook.com/powershell/ `
        -Credential $Credential `
        -Authentication Basic `
        -AllowRedirection `
        -ErrorAction SilentlyContinue
    if ( !($script:ExchangeLabsRS) -and !([String]::IsNullOrEmpty($error[0])) ) {
        Write-Error "Could not set up a remote connection to Exchange Labs:  $($error[0])"
        continue
    } else {
        Import-PSSession -Session $ExchangeLabsRS -AllowClobber
    }
    Write-Host "Succeeded."
} # end 'BEGIN{}'

# This section executes for each object in the pipeline.
PROCESS {
    if ( !($_) -and !($User) ) { return; }
    
    if ($_) { $User = $_ }
    
    # Was a username passed to us?  If not, bail.
    if (!($User)) { 
        Write-Error "USAGE:  New-OutlookLiveMailbox -User $User -Credential $Credential"
    }
    
    # Try to find the user in AD.
    if ($User.GetType().Name -eq "ArsUserObject") {
        # A ArsUserObject was passed; if '-f' was not specified, re-check
        # that it is a valid eID.
        if (!($f)) {
            $objUser = (Get-QADUser $User.Name).DirectoryEntry
        } else {
            $objUser = ($User).DirectoryEntry
        }
    } elseif ($User.GetType().Name -eq "String") {
        # A bare string was passed.  Turn it into a DirectoryEntry.
        $objUser = (Get-QADUser $User).DirectoryEntry
    } else {
        # We have no idea what this is.  Bail.
        Write-Host "$User is not a ArsUserObject or a String.  Cannot create mailbox."
        return
    }
    
    if (!($objUser)) {
        if ($f) {
            Write-Host "$User is not a valid user in Active Directory."
            Write-Host "Creating mailbox for non-existent user anyway ('-f' specified)..."
            $objUser = New-Object -TypeName System.DirectoryServices.DirectoryEntry
            $objUser.cn = $User
            $objUser.displayName = $User
            $objUser.givenName = "Unspecified"
            $objUser.sn = "Unspecified"
        } else {
            Write-Error "$User is not a valid user in Active Directory."
            return
        }
    } else {
        Write-Host "Found user $($objUser.cn) ($($objUser.displayName)) in Active Directory."
    }
    
    # Now make sure that the user doesn't already have a mailbox in Exchange Labs.
    # The easiest way is to trap the error produced if the user doesn't exist.
    $cmdGetMailbox = "Invoke-Command { Get-Mailbox $($objUser.cn) } -Session `$ExchangeLabsRS -ErrorAction SilentlyContinue"
    
    Write-Host -NoNewline "Checking that a mailbox for this user doesn't already exist..."
    $error.Clear()
    $mailbox = Invoke-Expression($cmdGetMailbox)
    
    #TODO:  This is ugly.  If $cmdGetMailbox returns an error, it will be 
    # contained in $mailbox, which falsely executes this block.
    # Check for an error condition instead (or in addition to).
    if ($mailbox) {
        Write-Error "A mailbox already exists for user $($objUser.cn)."
        return
    } else {
        Write-Host "no."
    }
    
    Write-Host -NoNewline "Creating an Exchange Labs mailbox for user $($objUser.cn)..."
    $Password = ConvertTo-SecureString 'P2ssw0rd!' -AsPlainText -Force
    $cmdNewMailbox = "New-Mailbox "                                   + `
        "-Name          `"$($objUser.cn)`" "                          + `
        "-WindowsLiveId '$($objUser.cn)@elr3.jmu.edu' "               + `
        "-DisplayName   `"$($objUser.sn), $($objUser.givenName)`" "   + `
        "-FirstName     `"$($objUser.givenName)`" "                   + `
        "-LastName      `"$($objUser.sn)`" "                          + `
        "-Password       `$Password "                                 + `
        "-ErrorAction SilentlyContinue"
    
    $error.Clear()

    # Run the command to create the mailbox.
    $mailbox = Invoke-Expression($cmdNewMailbox)

    if ($mailbox -and [String]::IsNullOrEmpty($error[0])) { 
        Write-Host "done." 
        Write-Host "Finished processing user $($objUser.cn)."
    } else {
        Write-Error "Mailbox creation was unsuccessful: $($error[0])"
    }
    
    Write-Host # Output a blank line
} # end 'PROCESS{}'

# This section executes only once, after the pipeline.    
END {
    # Close the connection to Exchange Labs.
    Write-Host -NoNewline "Closing connection to Exchange Labs..."
    Remove-PSSession -Session $ExchangeLabsRS
    Write-Host "done."

} # end 'END{}'
} # end function

Write-Host "Added New-OutlookLiveMailbox to global functions." -Fore White
