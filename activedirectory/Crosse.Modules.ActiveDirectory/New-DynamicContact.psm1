################################################################################
# 
# $Id$
#
# DESCRIPTION: This script will create a new dynamic contact object in
#              in Active Directory.  Please see 
#              http://www.ietf.org/rfc/rfc2589.txt for more information.
#              Note:  The target domain MUST be in a forest operating at the 
#              Windows 2003 forest functional level!
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


function New-DynamicContact {
    param (
            [string]
            # The email address to assign to the dynamic contact.
            $EmailAddress='',

            [string]
            # The Organizational Unit in which to place the dynamic contact.
            $OrganizationalUnit='',

            [string]
            # To specify the fully qualified domain name (FQDN) of the domain 
            # controller on which to perform the create.
            $DomainController='',

            [string]
            # The descrtiption of the contact.
            $Description='',

            [int]
            # The dynamic contact's Time-To-Live (TTL), in seconds.
            # The default is 900 seconds, or 15 minutes.  This is the lowest
            # value allowed by the Active Directory Schema.
            $EntryTTL=900,

            [switch]
            # Whether to print verbose information. By default this parameter
            # is set to $True.
            $Verbose=$True,

            $inputObject=$Null
        )

    BEGIN {
        # This has something to do with pipelining.  
        # Let's call it "magic voodoo" for now.
        if ($inputObject) {
            Write-Output $inputObject | &($MyInvocation.InvocationName) -EmailAddress $EmailAddress 
            break
        }
    }
    PROCESS {
        # If we got data via the pipeline, assign it to a named variable 
        # to make things easier to read.
        if ($_) { $EmailAddress = $_ }

        # Validate input.
        if ( !($EmailAddress) ) {
            Write-Error "The Email Address must be specified."
            return
        }

        if ( $OrganizationalUnit.IndexOfAny("/") -ge 0 ) {
            Write-Error "Invalid Organizational Unit"
            return
        }

        $SanitizedEmailAddress = $EmailAddress.Replace("@", "_at_")

        # Get the current domain for use later.
        $currDomain = [System.DirectoryServices.ActiveDirectory.Domain]::GetCurrentDomain()

        if ( !($OrganizationalUnit) ) {
            # Default to the Users container if no OU was specified.
            $OrganizationalUnit = "CN=Users"
        }

        if ( $OrganizationalUnit.ToLower().IndexOf("dc=") -lt 0 ) {
            # If the user left of the "DC=" part of the OU, get the
            # current domain and use that.
            $OrganizationalUnit += ",DC=$($currDomain.Name.Replace('.', ',DC='))"
        }

        # Create an ADSI connection.
        $ldap = "LDAP://"
        if ( $DomainController ) {
            $ldap += $DomainController + "/"
        }
        
        # Do this here just because it's easier to construct the URI.
        $contactExists = [System.DirectoryServices.DirectoryEntry]::Exists(
                "$($ldap)CN=$($SanitizedEmailAddress),$($OrganizationalUnit)")

        $ldap += $OrganizationalUnit
        $parentOU = [ADSI]$ldap

        $objUser = $null
        # Verify whether the object already exists.
        if ($contactExists) {
            $error.Clear()
            $ErrorActionPreference = "SilentlyContinue"

            # The contact already exists; just set the TTL
            $objUser = $parentOU.psbase.Children.Find("CN=$($SanitizedEmailAddress)")

            if ( !([String]::IsNullOrEmpty($error[0])) ) {
                Write-Error "Could not find contact, but supposedly it exists: $($error[0])"
                return
            }

            $objUser.Put("entryTTL", $entryTTL)

            # Commit the changes to Active Directory.
            $error.Clear()
            $objUser.SetInfo()
            $ErrorActionPreference = "Continue"

            if ( !([String]::IsNullOrEmpty($error[0])) ) {
                Write-Error "Could not modify contact $EmailAddress: $($error[0])"
            } else {
                if ($Verbose) {
                    Write-Host "Modified contact $EmailAddress"
                } else {
                    Write-Host -NoNewLine "!"
                }
            }
        } else {
            # Create the user and set some info.
            $objUser = $parentOU.Create("contact", "CN=$($SanitizedEmailAddress)")
            $objUser.Put("objectClass", @('dynamicObject', 'contact'))
            $objUser.Put("displayName", "$EmailAddress")
            $objUser.Put("name", "$EmailAddress")
            $objUser.Put("mail", "$EmailAddress")
            $objUser.Put("mailNickname", "$SanitizedEmailAddress")
            $objUser.Put("entryTTL", $entryTTL)
            $objUser.Put("proxyAddresses", @("SMTP:$($EmailAddress)", "smtp:$($SanitizedEmailAddress)@$($currDomain.Name)"))
            $objUser.Put("targetAddress", "SMTP:$($EmailAddress)")
            $objUser.Put("msExchHideFromAddressLists", "TRUE")
            $objUser.Put("msExchRequireAuthToSendTo", "TRUE")
            if ($Description) {
                $objUser.Put("description", "$($Description)")
            }

            # Commit the changes to Active Directory.
            $error.Clear()
            $ErrorActionPreference = "SilentlyContinue"
            $objUser.SetInfo()
            $ErrorActionPreference = "Continue"

            $objUser.Put("authOrig", @("CN=$($SanitizedEmailAddress),$($OrganizationalUnit)"))
            $objUser.SetInfo()

            if ( !([String]::IsNullOrEmpty($error[0])) ) {
                Write-Error "Could not create contact: $($error[0])"
            } else {
                if ($Verbose) {
                    Write-Host "Created contact $EmailAddress"
                } else {
                    Write-Host -NoNewLine "."
                }
            }
        }
    }
    END {
    }

    <#
        .SYNOPSIS
        Creates a "dynamic" contact in Active Directory

        .DESCRIPTION
        Creates a "dynamic" contact in Active Directory, as per RFC 2589.
        See "Related Links" for more information.

        .INPUTS
        None.  New-DynamicContact does not accept any values from the pipeline.

        .OUTPUTS
        None.  New-DynamicContact does not return any values.

        .LINK
        http://www.ietf.org/rfc/rfc2589.txt
        http://msdn.microsoft.com/en-us/library/ms677963%28VS.85%29.aspx
#>
}

