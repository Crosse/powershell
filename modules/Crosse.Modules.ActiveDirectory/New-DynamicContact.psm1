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
    [CmdletBinding(SupportsShouldProcess=$true,
            ConfirmImpact="High")]
    param (
            [Parameter(Mandatory=$true,
                ValueFromPipeline=$true)]
            [ValidateNotNullOrEmpty()]
            [string]
            # The CN to assign to the dynamic contact.  If not specified,
            # the default value will be a sanitized version of the email
            # address.
            $CN,

            [Parameter(Mandatory=$false)]
            [ValidatePattern(".*@*.\.*")]
            [string]
            # The email address to assign to the dynamic contact.
            $EmailAddress,

            [Parameter(Mandatory=$false)]
            [ValidateNotNullOrEmpty()]
            [string]
            # The contact's Display Name (displayName).
            $DisplayName,

            [Parameter(Mandatory=$false)]
            [ValidateNotNullOrEmpty()]
            [string]
            # The contact's Simple Display Name (displayNamePrintable).
            $SimpleDisplayName,

            [Parameter(Mandatory=$false)]
            [ValidateNotNullOrEmpty()]
            [string]
            # The contact's name.
            $Name,

            [Parameter(Mandatory=$false)]
            [ValidateNotNullOrEmpty()]
            [string]
            # The contact's first name (givenName).
            $FirstName,

            [Parameter(Mandatory=$false)]
            [ValidateNotNullOrEmpty()]
            [string]
            # The contact's last name (sn).
            $LastName,

            [Parameter(Mandatory=$false)]
            [switch]
            # The HiddenFromAddressListsEnabled parameter specifies whether 
            # this mailbox is hidden from other address lists.  The default
            # value is false.
            $HiddenFromAddressListsEnabled=$false,

            [Parameter(Mandatory=$false)]
            [switch]
            # The RequireSenderAuthenticationEnabled parameter specifies 
            # whether senders must be authenticated. The default value is 
            # false.
            $RequireSenderAuthenticationEnabled=$false,

            [Parameter(Mandatory=$false)]
            [ValidatePattern("(CN|OU)=.*")]
            [string]
            # The Organizational Unit in which to place the dynamic contact.
            # The default is "CN=Users". If the value is not fully-qualified
            # (i.e., if the value does not end with "DC=foo"), then the
            # current domain will be appended.
            $OrganizationalUnit='CN=Users',

            [Parameter(Mandatory=$false)]
            [ValidateNotNullOrEmpty()]
            [string]
            # The fully qualfied domain name (FQDN) of the domain 
            # controller on which to perform the create.
            $DomainController,

            [Parameter(Mandatory=$false)]
            [ValidateNotNullOrEmpty()]
            [string]
            # The descrtiption of the contact.
            $Description,

            [Parameter(Mandatory=$false)]
            [ValidateRange(900,31557600)]
            [int]
            # The dynamic contact's Time-To-Live (TTL), in seconds.
            # The default is 900 seconds, or 15 minutes.  This is the lowest
            # value allowed by the Active Directory Schema.
            $EntryTTL=900
        )

    BEGIN {
        if ( $OrganizationalUnit.ToLower().IndexOf("dc=") -lt 0 ) {
            # If the user left of the "DC=" part of the OU, get the
            # current domain and use that.
            $currDomain = [System.DirectoryServices.ActiveDirectory.Domain]::GetCurrentDomain()
            $domainDN = $currDomain.GetDirectoryEntry().distinguishedName
            $OrganizationalUnit += ",$($domainDN)"
            Write-Verbose "Using `"$OrganizationalUnit`" as the fully-qualified Organizational Unit"
        }

        $ou = $OrganizationalUnit

        # Create an ADSI connection.
        $ldapPrefix = "LDAP://"
        if ( $DomainController ) {
            $ldapPrefix += $DomainController + "/"
        }

        $error.Clear()
        $ldap = "$($ldapPrefix)$($OrganizationalUnit)"
        $parentOU = [ADSI]$ldap

        if ($parentOU.Path -eq $null) {
            Write-Error "Could not bind to `"$ldap`": $($error[0])"
            return $null
        }
    }
    PROCESS {
        if ([String]::IsNullOrEmpty($CN)) {
            $CN = $EmailAddress.Replace("@", "_at_")
        }

        $dn = "$($ldapPrefix)CN=$($CN),$($ou)"
        Write-Verbose "Searching for `"$dn`""

        $contactExists = [System.DirectoryServices.DirectoryEntry]::Exists($dn)

        $objUser = $null
        # Verify whether the object already exists.
        if ($contactExists) {
            Write-Verbose "Contact already exists."
            $error.Clear()
            $ErrorActionPreference = "SilentlyContinue"

            $objUser = $parentOU.psbase.Children.Find("CN=$($CN)")
            if ([String]::IsNullOrEmpty($EmailAddress)) {
                $EmailAddress = $objUser.Get("mail")
            }

            if ( !([String]::IsNullOrEmpty($error[0])) ) {
                Write-Error "Could not find contact, but supposedly it exists: $($error[0])"
                return
            }
        } else {
            # Create the user and set some info.
            Write-Verbose "Creating contact $CN"
            $error.Clear()
            $objUser = $parentOU.Create("contact", "CN=$($CN)")
            if ($objUser -eq $null) {
                Write-Error "Could not create dynamic object:  $($error[0])"
                return $null
            }

            $objUser.Put("objectClass", @('dynamicObject', 'contact'))
            $objUser.Put("mail", "$EmailAddress")
            $objUser.Put("mailNickname", "$CN")
            $objUser.Put("proxyAddresses", @("SMTP:$($EmailAddress)"))
            $objUser.Put("targetAddress", "SMTP:$($EmailAddress)")

            if (![String]::IsNullOrEmpty($Name)) {
                $objUser.Put("name", "$Name")
            }
            
            if (![String]::IsNullOrEmpty($DisplayName)) {
                $objUser.Put("displayName", "$DisplayName")
            }
            
            if (![String]::IsNullOrEmpty($SimpleDisplayName)) {
                $objUser.Put("displayNamePrintable", "$SimpleDisplayName")
            }
            
            if (![String]::IsNullOrEmpty($FirstName)) {
                $objUser.Put("givenName", $FirstName)
            }

            if (![String]::IsNullOrEmpty($LastName)) {
                $objUser.Put("sn", $LastName)
            }
            
            if (![String]::IsNullOrEmpty($Description)) {
                $objUser.Put("description", "$($Description)")
            }

            if ($HiddenFromAddressListsEnabled -eq $true) {
                $objUser.Put("msExchHideFromAddressLists", $true)
            }

            if ($RequireSenderAuthenticationEnabled -eq $true) {
                $objUser.Put("msExchRequireAuthToSendTo", $true)
            }
        }

        $objUser.Put("entryTTL", $entryTTL)
        # Commit the changes to Active Directory.
        Write-Verbose "Committing the changes to Active Directory"
        $error.Clear()
        $ErrorActionPreference = "SilentlyContinue"
        $objUser.SetInfo()
        $ErrorActionPreference = "Continue"

        # The following line is the Exchange equivalent of "Allowed Senders"
        # $objUser.Put("authOrig", @("CN=$($CN),$($ou)"))
        # $objUser.SetInfo()

        if (![String]::IsNullOrEmpty($error[0])) {
            Write-Error "Could not create or update contact: $($error[0])"
        } else {
            if ($contactExists) {
                Write-Verbose "Updated contact $CN"
            } else {
                Write-Verbose "Created contact $CN"
            }
        }
        return $objUser
    }

    <#
        .SYNOPSIS
        Creates a dynamic contact in Active Directory

        .DESCRIPTION
        Creates a dynamic contact in Active Directory, as per RFC 2589.
        See "Related Links" for more information.

        .INPUTS
        New-DynamicContact can accept the a System.String from the pipeline that 
        corresponds to the email address to use for the new contact.

        .OUTPUTS
        System.DirectoryServices.DirectoryEntry. New-DynamicContact returns the 
        newly-created (or updated) dynamic object, or $null if an error 
        occurred.

        .EXAMPLE
        C:\PS> New-DynamicContact user@foo.com


        distinguishedName : {CN=user_at_foo.com,CN=Users,DC=contoso,DC=com}
        Path              : LDAP://CN=user_at_foo.com,CN=Users,DC=contoso,DC=com

        The above example creates a new dynamic contact with default parameters.

        .EXAMPLE

        C:\PS> New-DynamicContact -Verbose -EmailAddress "asdf@asdf.com" -OrganizationalUnit "OU=DynamicObjects,OU=Test" -Description "test" -EntryTTL 1000
        VERBOSE: Using "OU=DynamicObjects,OU=Test,DC=contoso,DC=com" as the fully-qualified Organizational Unit
        VERBOSE: The requested object already exists.
        VERBOSE: Updating TTL value for already-existing object
        VERBOSE: Committing changes to Active Directory
        VERBOSE: Changes committed.


        distinguishedName : {CN=asdf_at_asdf.com,OU=DynamicObjects,OU=Test,DC=contoso,DC=com}
        Path              : LDAP://CN=asdf_at_asdf.com,OU=DynamicObjects,OU=Test,DC=contoso,DC=com

        The above exmaple creates a new dynamic contact with data specified on the command-line.

        .LINK
        http://www.ietf.org/rfc/rfc2589.txt

        .LINK
        http://msdn.microsoft.com/en-us/library/ms677963%28VS.85%29.aspx
#>
}

