################################################################################
#
# Copyright (c) 2009 - 2011 Seth Wright <wrightst@jmu.edu>
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


################################################################################
<#
    .SYNOPSIS
    Retrieves attributes from Active Directory for the specified object.

    .DESCRIPTION
    Retrieves attributes from Active Directory for the specified object. The
    object specified by the "Identity" parameter is searched for first as a user,
    then as a computer object.

    .INPUTS
    System.String.  The Identity (or Identities) for which to retrieve attributes
    can be passed via the command line.

    .OUTPUTS
    A PSObject with the requested attributes for the Identity.

    .EXAMPLE
    PS C:\> Get-ADAttribute -Identity wrightst -Attributes mail | Format-List

    Name              : wrightst
    DistinguishedName : CN=wrightst,OU=Users,...
    mail              : wrightst@jmu.edu

    The above example illustrates how to retrieve attributes for a user.

    .EXAMPLE
    PS C:\> Get-ADAttribute -Identity tex-vm -Attributes lastLogonTimestamp,pwdLastSet | Format-List

    Name               : tex-vm
    DistinguishedName  : CN=TEX-VM,OU=Desktops,...
    lastLogonTimestamp : 10/14/2011 12:51:44 PM
    pwdLastSet         : 9/14/2011 9:46:12 PM

    The above example illustrates how to retrieve attributes for a computer.
    Notice is is the same as for a user.  Also, notice that attributes that should
    be datetimes are returned as such instead of in the Int64 or LargeInteger
    format.
#>
################################################################################
function Get-ADAttribute {
    [CmdletBinding(SupportsShouldProcess=$true,
            ConfirmImpact="High")]
    param (
            [Parameter(Mandatory=$true,
                ValueFromPipeline=$true)]
            [ValidateNotNullOrEmpty()]
            [string]
            # Specifies the object for which to search.
            $Identity,

            [Parameter(Mandatory=$true)]
            [string[]]
            # Specifies which attributes to return.
            $Attributes,

            [Parameter(Mandatory=$false)]
            [string]
            # The domain controller to target.
            $DomainController
        )

    BEGIN {
        $schema = [System.DirectoryServices.ActiveDirectory.ActiveDirectorySchema]::GetCurrentSchema()
        if ([String]::IsNullOrEmpty($DomainController)) {
            $pc = New-Object System.DirectoryServices.AccountManagement.PrincipalContext Domain
        } else {
            $pc = New-Object System.DirectoryServices.AccountManagement.PrincipalContext Domain, $DomainController
        }

        $int64Attributes = New-Object System.Collections.ArrayList
        foreach ($attribute in $Attributes) {
            if ($schema.FindProperty($attribute).Syntax -eq "Int64") {
                $null = $int64Attributes.add($attribute)
            }
        }
    }
    PROCESS {
        $props = New-Object PSObject -Property @{
            Name                = $Identity
            DistinguishedName   = $null
        }

        $objUser=
            [System.DirectoryServices.AccountManagement.UserPrincipal]::FindByIdentity($pc, $Identity)

        if ($objUser -eq $null) {
            $objUser = [System.DirectoryServices.AccountManagement.ComputerPrincipal]::FindByIdentity($pc, $Identity)
            if ($objUser -eq $null) {
                Write-Error "Cannot find object $Identity in the current Active Directory domain."
                return
            }
        }

        $props.DistinguishedName = $objUser.DistinguishedName
        Write-Verbose "Found $($objUser.StructuralObjectClass) DN: $($objUser.DistinguishedName)"

        $dirEntry = $objUser.GetUnderlyingObject()
        foreach ($attribute in $Attributes) {
            $prop = $dirEntry.InvokeGet($attribute)
            if ($prop -ne $null -and $int64Attributes.Contains($attribute)) {
                Write-Verbose "Converting attribute $attribute to Int64"
                $prop = $dirEntry.ConvertLargeIntegerToInt64($prop)
                if ($attribute -match 'date' -or
                        $attribute -match 'time' -or
                        $attribute -match 'lastSet') {
                    $prop = [DateTime]::FromFileTime($prop)
                }
            } elseif ($prop -ne $null -and $attribute -match 'Guid') {
                $prop = [Guid]$prop
            }

            $props | Add-Member NoteProperty $attribute $prop
            Write-Verbose ("{0} = {1}" -f $Attribute, $props.$attribute)
        }

        $props
    }
}
