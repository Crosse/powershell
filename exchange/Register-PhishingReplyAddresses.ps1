################################################################################
# 
# $URL$
# $Author$
# $Date$
# $Rev$
#
# DESCRIPTION:  Imports a phishing_reply_addresses file of the format found at 
#               http://aper.svn.sourceforge.net/viewvc/aper/phishing_reply_addresses
#               into Active Directory.
#               This script can be paired with the New-DynamicContact script
#               to create dynamic objects, or with other cmdlets that create
#               users.  By default it uses New-DynamicContact.
#
# Redistribution and use in source and binary forms, with or without           
# modification, are permitted provided that the following conditions are met:  
#
#  1. Redistributions of source code must retain the above copyright notice,   
#     this list of conditions and the following disclaimer.                    
#  2. Redistributions in binary form must reproduce the above copyright        
#     notice, this list of conditions and the following disclaimer in the      
#     documentation and/or other materials provided with the distribution.     
# 
# THIS SOFTWARE IS PROVIDED BY THE AUTHOR AND CONTRIBUTORS ``AS IS'' AND ANY   
# EXPRESS OR IMPLIED WARRANTIES, INCLUDING, BUT NOT LIMITED TO, THE IMPLIED    
# WARRANTIES OF MERCHANTABILITY AND FITNESS FOR A PARTICULAR PURPOSE ARE       
# DISCLAIMED. IN NO EVENT SHALL THE AUTHOR OR CONTRIBUTORS BE LIABLE FOR ANY   
# DIRECT, INDIRECT, INCIDENTAL, SPECIAL, EXEMPLARY, OR CONSEQUENTIAL DAMAGES   
# (INCLUDING, BUT NOT LIMITED TO, PROCUREMENT OF SUBSTITUTE GOODS OR SERVICES; 
# LOSS OF USE, DATA, OR PROFITS; OR BUSINESS INTERRUPTION) HOWEVER CAUSED AND  
# ON ANY THEORY OF LIABILITY, WHETHER IN CONTRACT, STRICT LIABILITY, OR TORT   
# (INCLUDING NEGLIGENCE OR OTHERWISE) ARISING IN ANY WAY OUT OF THE USE OF     
# THIS SOFTWARE, EVEN IF ADVISED OF THE POSSIBILITY OF SUCH DAMAGE.            
# 
################################################################################


param ([switch]$Install, [string]$Path='', [string]$OrganizationalUnit='', `
        [string]$DomainController='', [string]$EntryTTL="604800", [switch]$Verbose=$False)

# Note that the default EntryTTL is one week in seconds (7d * 24h * 60m * 60s)

if (Test-Path function:Register-PhishingReplyAddresses) {
    Remove-Item function:Register-PhishingReplyAddresses
}

function global:Register-PhishingReplyAddresses([string]$Path='', [string]$OrganizationalUnit='', `
                                                [string]$DomainController='', `
                                                [switch]$Verbose=$False, $inputObject=$Null) {
    BEGIN {
        # This has something to do with pipelining.  
        # Let's call it "magic voodoo" for now.
        if ($inputObject) {
            Write-Output $inputObject | &($MyInvocation.InvocationName) -Path $Path `
                                        -OrganizationalUnit $OrganizationalUnit `
                                        -DomainController $DomainController -Verbose:$Verbose
            break
        }

        # Bail if the OU isn't present, or is in the wrong format.
        if ([String]::IsNullOrEmpty($OrganizationalUnit)) {
            Write-Error "The Organizational Unit must be specified."
            exit
        }
        if ($OrganizationalUnit.IndexOfAny("/") -ge 0) {
            Write-Error "Invalid Organizational Unit.  Please specify in LDAP format"
            exit
        }

        # Set the default URL is no URL or file path is specified.
        if ([String]::IsNullOrEmpty($Path)) {
            Write-Warning "Using default URL for the anti-phishing-reply list"
            $Path = "http://aper.svn.sourceforge.net/viewvc/aper/phishing_reply_addresses"
        }
    }
    PROCESS {
        if ( !($Path) -and !($inputObject) ) {
            Write-Error "No file was specified."
            exit
        }

        $lines = $null

        # Attempt to download the file if a URL was specified.
        if ( $Path.StartsWith("http") ) {
            if ($Verbose) {
                Write-Host "Attempting to download the requested file..."
            }
            $wc = New-Object Net.WebClient
            $now = Get-Date
            # Save it in a file with today's date, then set the $Path to be this
            # local file for further processing.
            $fileName = "$($pwd)\phishing-reply-addresses_$($now.Year)-$($now.Month)-$($now.Day).txt"
            $wc.DownloadFile($Path, $fileName)
            $Path = $fileName
        } 

        if ( !(Test-Path $Path) ) {
            Write-Error "Path does not exist"
            exit
        }

        # Open the file and start processing.
        $lines = Get-Content $Path

        # Either the file was empty, or something else happened that 
        # prevented us from reading it.  Bail.
        if ( !($lines) ) {
            Write-Error "Could not read file"
            exit
        }

        if ($Verbose) {
            Write-Host "File contains $($lines.Count) lines"
        }

        # Get the current domain for use later.
        $currDomain = [System.DirectoryServices.ActiveDirectory.Domain]::GetCurrentDomain()

        # If the user left of the "DC=" part of the OU, get the
        # current domain and use that.
        if ( $OrganizationalUnit.ToLower().IndexOf("dc=") -lt 0 ) {
            $OrganizationalUnit += ",DC=$($currDomain.Name.Replace('.', ',DC='))"
        }

        # Create an ADSI connection.
        $preamble = "LDAP://"
        if ( $DomainController ) {
            $preamble += $preamble + $DomainController + "/"
        }
        $ldap = $preamble + $OrganizationalUnit
        $parentOU = [ADSI]$ldap

        # Iterate through all of the addresses in the specified file and
        # process them.
        $i = 0
        foreach ($line in $lines) {
            # Skip comment and empty lines.
            if ($line.StartsWith('#') -or $line.Length -eq 0) {
                continue
            }
            $error.Clear()
            $ErrorActionPreference = "SilentlyContinue"
            $address = $line.Remove($line.IndexOf(','))
            # If there was no comma in the line (the previous statement failed),
            # then report that a bad line was found.
            if (!([String]::IsNullOrEmpty($error[0]))) {
                Write-Host "Malformed line:  `"$line`""
            }
            $ErrorActionPreference = "Continue"

            # Do a dirty validation that $address contains something like an email address
            if ( !($address.Contains('@') )) {
                Write-Warning "Malformed email address:  `"$address`""
                continue
            }

            # TODO:  perform better sanitation on the purported email address.
            $SanitizedAddress = $address.Replace("@", "_at_")

            $objUser = $null
            # Verify whether the object already exists.
            $exists = [System.DirectoryServices.DirectoryEntry]::Exists(
                    "$($preamble)CN=$($SanitizedAddress),$($OrganizationalUnit)")
            if ($exists) {
                $error.Clear()
                $ErrorActionPreference = "SilentlyContinue"

                # The contact already exists; just set the TTL
                $objUser = $parentOU.psbase.Children.Find("CN=$($SanitizedAddress)")

                # Wait a minute...something's not right.  We found the object 
                # earlier, but for some reason we can't now.  Report this and move on.
                if ( !([String]::IsNullOrEmpty($error[0])) ) {
                    Write-Error "Could not find contact, but supposedly it exists: $($error[0])"
                    continue
                }

                $objUser.Put("entryTTL", $entryTTL)
                $objUser.Put("description", "anti-phishing-reply ($line)")
                # The following line clears the authOrig attribute.
                # We're doing this so that we can apply transport rules to
                # the phishing contacts, instead of the sender only receiving
                # an NDR.
                # ADS_PROPERTY_DELETE = 1; the '0' is not well explained.
                $objUser.PutEx(1, "authOrig", 0)

                # Commit the changes to Active Directory.
                $error.Clear()
                $objUser.SetInfo()
                $ErrorActionPreference = "Continue"

                if ( !([String]::IsNullOrEmpty($error[0])) ) {
                    Write-Error "Could not modify contact $address: $($error[0])"
                } else {
                    if ($Verbose) {
                        Write-Host "Modified contact $address"
                    }
                }
            } else {
                # Create the user and set some info.
                $objUser = $parentOU.Create("contact", "CN=$($SanitizedAddress)")
                $objUser.Put("objectClass", @('dynamicObject', 'contact'))
                $objUser.Put("displayName", "$address")
                $objUser.Put("name", "$address")
                $objUser.Put("mail", "$address")
                $objUser.Put("mailNickname", "$SanitizedAddress")
                $objUser.Put("entryTTL", $entryTTL)
                $objUser.Put("proxyAddresses", @("SMTP:$($address)", "smtp:$($SanitizedAddress)@$($currDomain.Name)"))
                $objUser.Put("targetAddress", "SMTP:$($address)")
                $objUser.Put("msExchHideFromAddressLists", "TRUE")
                $objUser.Put("msExchRequireAuthToSendTo", "TRUE")
                $objUser.Put("description", "anti-phishing-reply ($line)")

                # Commit the changes to Active Directory.
                $error.Clear()
                $ErrorActionPreference = "SilentlyContinue"
                $objUser.SetInfo()
                $ErrorActionPreference = "Continue"

                #$objUser.Put("authOrig", @("CN=$($SanitizedAddress),$($OrganizationalUnit)"))
                #$objUser.SetInfo()

                if ( !([String]::IsNullOrEmpty($error[0])) ) {
                    Write-Error "Could not create contact: $($error[0])"
                } else {
                    if ($Verbose) {
                        Write-Host "Created contact $address"
                    }
                }
            }

            $i++
            if ( ($i % 10) -eq 0) {
                Write-Progress -Activity "Processing $($lines.Count) Lines..." -Status "$([int]($i/$lines.Count*100))% Complete" -percentComplete ($i/$lines.Count*100) -CurrentOperation "$address"
            }
        }
        Write-Host "`nFinished loading anti-phishing-reply list."
    }
    END {
    }
}

if ($Install -eq $True) {
    Write-Host "Added Register-PhishingReplyAddresses to global functions." -Fore White
    exit
} else {
    $now = Get-Date
    Start-Transcript "anti-phishing-reply_$($now.Year)-$($now.Month)-$($now.Day).log"
    Register-PhishingReplyAddresses -Path $Path -Verbose:$Verbose -OrganizationalUnit $OrganizationalUnit -DomainController $DomainController
    Stop-Transcript
}
