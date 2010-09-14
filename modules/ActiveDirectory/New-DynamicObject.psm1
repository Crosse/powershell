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


function New-DynamicObject {
    [CmdletBinding(SupportsShouldProcess=$true,
            ConfirmImpact="High")]
    param (
            [Parameter(Mandatory=$true,
                ParameterSetName="User")]
            [switch]
            # Specifies that the object should be created as a dynamic 
            # user object.
            $User,

            [Parameter(Mandatory=$true,
                ParameterSetName="Contact")]
            [switch]
            # Specifies that the object should be created as a dynamic 
            # contact object.
            $Contact,

            [switch]
            # Specifies whether to create the object with some basic 
            # attributes used for Microsoft Exchange.  The default value is 
            # false.
            $ExchangeObject=$false,

            [Parameter(Mandatory=$true,
                ValueFromPipeline=$true)]
            [ValidateNotNullOrEmpty()]
            [string]
            # The name to assign to the dynamic object.  This will be the LDAP
            # cn and 'name' attributes.
            $Name,

            [Parameter(ParameterSetName="User")]
            [ValidateLength(1,20)]
            [string]
            # The object's sAmAccountName.  The default value is a sanitized
            # version of the Name attribute.
            $SamAccountName,

            [Parameter(ParameterSetName="User")]
            [ValidatePattern(".*@*.\.*")]
            [string]
            # The object's userPrincipalName.  The default value is the Name
            # attribute followed by "@current.domain.com"
            $UserPrincipalName,

            [ValidateNotNullOrEmpty()]
            [string]
            # The object's Exchange Alias (mailNickname).  The default value
            # is a sanitized version of the 'name' attribute.
            $Alias,

            [Parameter(Mandatory=$true,
                    ParameterSetName="Contact")]
            [Parameter(Mandatory=$false,
                    ParameterSetName="User")]
            [ValidatePattern(".*@*.\.*")]
            [string]
            # The email address to assign to the dynamic object.
            $EmailAddress,

            [ValidateNotNullOrEmpty()]
            [string]
            # The object's Display Name (displayName).
            $DisplayName,

            [ValidateNotNullOrEmpty()]
            [string]
            # The object's Simple Display Name (displayNamePrintable).
            $SimpleDisplayName,

            [ValidateNotNullOrEmpty()]
            [string]
            # The object's first name (givenName).
            $FirstName,

            [ValidateNotNullOrEmpty()]
            [string]
            # The object's last name (sn).
            $LastName,

            [ValidateNotNullOrEmpty()]
            [string]
            # The description of the object.
            $Description,

            [switch]
            # The HiddenFromAddressListsEnabled parameter specifies whether 
            # this object is hidden from other address lists.  The default
            # value is false.
            $HiddenFromAddressListsEnabled=$false,

            [switch]
            # The RequireSenderAuthenticationEnabled parameter specifies 
            # whether senders must be authenticated. The default value is 
            # false.
            $RequireSenderAuthenticationEnabled=$false,

            [Parameter(ParameterSetName="User")]
            [ValidateNotNullOrEmpty()]
            # A SecureString representation of the password to use for the 
            # new user. The default is to create a user with no password set, 
            # thus creating the user in a default-disabled state.
            [System.Security.SecureString]
            $Password,

            [Parameter(ParameterSetName="User")]
            [switch]
            # If the cmdlet should prompt for a password to be entered instead
            # of passing the password on the command-line.
            $PromptForPassword=$false,

            [ValidatePattern("(CN|OU)=.*")]
            [string]
            # The Organizational Unit in which to place the dynamic object.
            # The default is "CN=Users". If the value is not fully-qualified
            # (i.e., if the value does not end with "DC=foo"), then the
            # current domain will be appended.
            $OrganizationalUnit='CN=Users',

            [ValidateNotNullOrEmpty()]
            [string]
            # The fully qualfied domain name (FQDN) of the domain 
            # controller on which to perform the create.
            $DomainController,

            [ValidateRange(900,31557600)]
            [int]
            # The dynamic object's Time-To-Live (TTL), in seconds.
            # The default is 900 seconds, or 15 minutes.  This is the lowest
            # value allowed by the Active Directory Schema.
            $EntryTTL=900
        )

    BEGIN {
        if (!$ExchangeObject -and 
                ($HiddenFromAddressListsEnabled -or 
                 $RequireSenderAuthenticationEnabled)) {
            Write-Error "The -ExchangeObject parameter must be used to specify -HiddenFromAddressListsEnabled or -RequireSenderAuthenticationEnabled"
            return $null
        }

        $currDomain = [System.DirectoryServices.ActiveDirectory.Domain]::GetCurrentDomain()
        if ($OrganizationalUnit.ToLower().IndexOf("dc=") -lt 0) {
            # If the user left of the "DC=" part of the OU, get the
            # current domain and use that.
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
            Write-Error "Could not bind to `"$ldap`".  Check for errors in the path."
            return $null
        }
    }
    PROCESS {
        $dn = "$($ldapPrefix)CN=$($Name),$($ou)"
        Write-Verbose "Searching for `"$dn`""

        $objectExists = [System.DirectoryServices.DirectoryEntry]::Exists($dn)

        $objUser = $null
        # Verify whether the object already exists.
        if ($objectExists) {
            Write-Warning "An object with the same DN already exists ($dn)."

            if (!$PSCmdlet.ShouldProcess($dn, "update Entry Time-To-Live")) {
                return $null
            }

            $error.Clear()
            $ErrorActionPreference = "SilentlyContinue"

            $objUser = $parentOU.psbase.Children.Find("CN=$($Name)")

            if ( !([String]::IsNullOrEmpty($error[0])) ) {
                Write-Error "Could not find contact, but supposedly it exists: $($error[0])"
                return $null
            }
        } else {
            if ($User -and $PromptForPassword) {
                $Password = Read-Host -AsSecureString -Prompt "New Password"
            }

            if (!$PSCmdlet.ShouldProcess($dn, "create object")) {
                return $null
            }
            
            # Create the user and set some info.
            Write-Verbose "Creating contact $dn"
            $error.Clear()
            $objUser = $parentOU.Create("contact", "CN=$($Name)")
            if ($objUser -eq $null) {
                Write-Error "Could not create dynamic object:  $($error[0])"
                return $null
            }

            if ($Contact) {
                $objUser.Put("objectClass", @('dynamicObject', 'contact'))
            } elseif ($User) {
                $objUser.Put("objectClass", @('dynamicObject', 'user'))

                if ([String]::IsNullOrEmpty($SamAccountName)) {
                    $SamAccountName = $Name.Replace(" ", "")
                }
                $objUser.Put("sAmAccountName", "$SamAccountName")

                if ([String]::IsNullOrEmpty($UserPrincipalName)) {
                    $UserPrincipalName = $SamAccountName + "@" + $currDomain.Name
                }
                $objUser.Put("userPrincipalName", "$UserPrincipalName")
            }

            if ($ExchangeObject) {
                if ([String]::IsNullOrEmpty($Alias)) {
                    $Alias = $Name.Replace(" ", "")
                }
                $objUser.Put("mailNickname", "$Alias")
            }

            if (![String]::IsNullOrEmpty($EmailAddress)) {
                $objUser.Put("mail", "$EmailAddress")
                if ($ExchangeObject) {
                    $objUser.Put("proxyAddresses", @("SMTP:$($EmailAddress)"))
                    $objUser.Put("targetAddress", "SMTP:$($EmailAddress)")
                }
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
        # $objUser.Put("authOrig", @("CN=$($Name),$($ou)"))

        if ($User -and ![String]::IsNullOrEmpty($Password)) {
            $bStr = [Runtime.InteropServices.Marshal]::SecureStringToBSTR($Password)
            $objUser.SetPassword([Runtime.InteropServices.Marshal]::PtrToStringAuto($bStr))
            [Runtime.InteropServices.Marshal]::ZeroFreeBSTR($bStr)

            $objUser.InvokeSet("AccountDisabled", $false)
            $objUser.InvokeSet("PasswordRequired", $true)

            # The line below works just as well as the two above, but I think
            # the lines above are more descriptive.
            # $objUser.Put("userAccountControl", 512)
            
            $objUser.SetInfo()
        }

        if (![String]::IsNullOrEmpty($error[0])) {
            Write-Error "Could not create or update object: $($error[0])"
        } else {
            if ($objectExists) {
                Write-Verbose "Updated object $dn"
            } else {
                Write-Verbose "Created object $dn"
            }
        }
        return $objUser
    }

    <#
        .SYNOPSIS
        Creates a dynamic object in Active Directory

        .DESCRIPTION
        Creates a dynamic object in Active Directory, as per RFC 2589.  See "Related Links" for more information.

        .INPUTS
        New-DynamicObject can accept the a System.String from the pipeline that corresponds to the name to use for the new object.

        .OUTPUTS
        System.DirectoryServices.DirectoryEntry. New-DynamicObject returns the newly-created (or updated) dynamic object, or $null if an error occurred.

        .EXAMPLE

        C:\PS> $passwd = Read-Host -AsSecureString
        *********
        C:\PS> New-DynamicObject -User -Name "TestUser1" -Password $passwd

        Confirm
        Are you sure you want to perform this action?
        Performing operation "create object" on Target "LDAP://CN=TestUser1,CN=Users,DC=contoso,DC=com".
        [Y] Yes  [A] Yes to All  [N] No  [L] No to All  [S] Suspend  [?] Help (default is "Y"):


        distinguishedName : {CN=TestUser1,CN=Users,DC=contoso,DC=com}
        Path              : LDAP://CN=TestUser1,CN=Users,DC=contoso,DC=com
        
        .EXAMPLE
        C:\PS> New-DynamicObject -Contact -Name "TestContact1" -EmailAddress "fdsa@asdf.com" -OrganizationalUnit "ou=DynamicObjects,ou=Test"
        WARNING: An object with the same DN already exists
        (LDAP://CN=TestContact1,ou=DynamicObjects,ou=Test,DC=contoso,DC=com).

        Confirm
        Are you sure you want to perform this action?
        Performing operation "update Entry Time-To-Live" on Target
        "LDAP://CN=TestContact1,ou=DynamicObjects,ou=Test,DC=contoso,DC=com".
        [Y] Yes  [A] Yes to All  [N] No  [L] No to All  [S] Suspend  [?] Help (default is "Y"):


        distinguishedName : {CN=TestContact1,OU=DynamicObjects,OU=Test,DC=contoso,DC=com}
        Path              : LDAP://CN=TestContact1,ou=DynamicObjects,ou=Test,DC=contoso,DC=com



        .LINK
        http://www.ietf.org/rfc/rfc2589.txt

        .LINK
        http://msdn.microsoft.com/en-us/library/ms677963%28VS.85%29.aspx
#>
}

