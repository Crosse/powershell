################################################################################
# 
# $Id$
#
# DESCRIPTION: This script will create a new dynamic user object in
#              in Active Directory.  Please see 
#              http://www.ietf.org/rfc/rfc2589.txt for more information.             
#              Note:  The target domain MUST be in a forest operating at the 
#              Windows 2003 forest functional level.
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


param ([string]$CN="", 
        [string]$OrganizationalUnit='', 
        [string]$DomainController='',
        [string]$SamId='', 
        [string]$UserPrincipalName='', 
        [int]$entryTTL=900,
        [string]$Description='', 
        [switch]$Install=$False)

# Note that the the default $entryTTL value above is the default value for the 
# DefaultMinTTL attribute in the Active Directory Schema.

if (Test-Path function:New-DynamicUser) {
    Remove-Item function:New-DynamicUser
}

function global:New-DynamicUser(
        [string]$CN='',
        [string]$OrganizationalUnit='',
        [string]$DomainController='', 
        [string]$SamId='',
        [string]$UserPrincipalName='',
        [string]$Description='', 
        [int]$EntryTTL=900,
        $inputObject=$Null) {
    BEGIN {
        # This has something to do with pipelining.  
        # Let's call it "magic voodoo" for now.
        if ($inputObject) {
            Write-Output $inputObject | &($MyInvocation.InvocationName) -CN $CN 
            break
        }
    }
    PROCESS {
        # If we got data via the pipeline, assign it to a named variable 
        # to make things easier to read.
        if ($_) { $CN = $_ }

        # Validate input.
        if ( !($CN) ) {
            Write-Error "The CN must be specified."
            return
        }

        if ( !($SamId) ) {
            # Set the sAMAccountName to the CN, if not specified.
            $SamId = $CN
        }

        # Get the current domain for use later.
        $currDomain = [System.DirectoryServices.ActiveDirectory.Domain]::GetCurrentDomain()

        if ( !($UserPrincipalName) ) {
            # Get the current domain's Name and use that for the UPN.
            $UserPrincipalName = "$($CN)@$($currDomain.Name)"
        }

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
        $ldap += $OrganizationalUnit
        $parentOU = [ADSI]$ldap

        # "Create the user and set some info.
        $objUser = $parentOU.Create("user", "CN=$($CN)")
        $objUser.Put("sAMAccountName", "$($SamId)")
        $objUser.Put("userPrincipalName", "$($UserPrincipalName)")
        if ($Description) {
            $objUser.Put("description", "$($Description)")
        }
        $objUser.Put("objectClass", @('dynamicObject', 'user'))
        $objUser.Put("entryTTL", $entryTTL)

        # Commit the changes to Active Directory.
        $error.Clear()
        $ErrorActionPreference = "SilentlyContinue"
        $objUser.SetInfo()
        $ErrorActionPreference = "Continue"

        if ( !([String]::IsNullOrEmpty($error[0])) ) {
            Write-Error "Could not create user: $($error[0])"
        } else {
            Write-Host "Created user $($objUser.Path)"
        }
    }
    END {
    }
}

if ($Install -eq $True) {
    Write-Host "Added New-DynamicUser to global functions." -Fore White
    return
} else {
    New-DynamicUser -CN $CN -OrganizationalUnit $OrganizationalUnit `
                    -DomainController $DomainController -SamId $SamId `
                    -UserPrincipalName $UserPrincipalName -Description $Description `
                    -entryTTL $entryTTL
}
