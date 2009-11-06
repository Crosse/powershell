################################################################################
# 
# $Id$
#
# DESCRIPTION: This script will create a new dynamic user object in
#              in Active Directory.  Please see 
#              http://www.ietf.org/rfc/rfc2589.txt for more information.             
#              Note:  The target domain MUST be in a forest operating at the 
#              Windows 2003 forest functional level!
# 
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
