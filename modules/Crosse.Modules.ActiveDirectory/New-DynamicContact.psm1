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
            [Parameter(Mandatory=$true,
                ValueFromPipeline=$true)]
            [ValidatePattern(".*@*.\.*")]
            [string]
            # The email address to assign to the dynamic contact.
            $EmailAddress,

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

    PROCESS {
        $SanitizedEmailAddress = $EmailAddress.Replace("@", "_at_")

        # Get the current domain for use later.
        $currDomain = [System.DirectoryServices.ActiveDirectory.Domain]::GetCurrentDomain()

        if ( $OrganizationalUnit.ToLower().IndexOf("dc=") -lt 0 ) {
            # If the user left of the "DC=" part of the OU, get the
            # current domain and use that.
            $OrganizationalUnit += ",DC=$($currDomain.Name.Replace('.', ',DC='))"
            Write-Verbose "Using `"$OrganizationalUnit`" as the fully-qualified Organizational Unit"
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

        if ($parentOU.Path -eq $null) {
            Write-Error "Could not bind to `"$ldap`": $($error[0])"
            return $null
        }

        $objUser = $null
        # Verify whether the object already exists.
        if ($contactExists) {
            Write-Verbose "The requested object already exists."

            $error.Clear()
            $ErrorActionPreference = "SilentlyContinue"

            # The contact already exists; just set the TTL
            Write-Verbose "Updating TTL value for already-existing object"
            $objUser = $parentOU.psbase.Children.Find("CN=$($SanitizedEmailAddress)")

            if ( !([String]::IsNullOrEmpty($error[0])) ) {
                Write-Error "Could not find contact, but supposedly it exists: $($error[0])"
                return
            }

            $objUser.Put("entryTTL", $entryTTL)

            # Commit the changes to Active Directory.
            Write-Verbose "Committing changes to Active Directory"
            $error.Clear()
            $objUser.SetInfo()
            $ErrorActionPreference = "Continue"

            if ( !([String]::IsNullOrEmpty($error[0])) ) {
                Write-Error "Could not modify contact $EmailAddress: $($error[0])"
            } else {
                Write-Verbose "Changes committed."
            }
        } else {
            # Create the user and set some info.
            Write-Verbose "Creating contact $SanitizedEmailAddress"
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
            if (![String]::IsNullOrEmpty($Description)) {
                $objUser.Put("description", "$($Description)")
            }

            # Commit the changes to Active Directory.
            Write-Verbose "Committing the changes to Active Directory"
            $error.Clear()
            $ErrorActionPreference = "SilentlyContinue"
            $objUser.SetInfo()
            $ErrorActionPreference = "Continue"

            $objUser.Put("authOrig", @("CN=$($SanitizedEmailAddress),$($OrganizationalUnit)"))
            $objUser.SetInfo()

            if ( !([String]::IsNullOrEmpty($error[0])) ) {
                Write-Error "Could not create contact: $($error[0])"
            } else {
                Write-Verbose "Created contact $EmailAddress"
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
        New-DynamicUser can accept the a System.String from the pipeline that 
        corresponds to the email address to use for the new contact.

        .OUTPUTS
        System.DirectoryServices.DirectoryEntry. New-DynamicUser returns the 
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
        http://msdn.microsoft.com/en-us/library/ms677963%28VS.85%29.aspx
#>
}

