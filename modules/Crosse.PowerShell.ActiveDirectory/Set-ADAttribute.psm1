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
    Modifies an attribute on an Active Directory object.

    .DESCRIPTION
    Modifies an attribute on an Active Directory object. This cmdlet must be run
    as a user with rights to modify the attribute in Active Directory.

    .INPUTS
    System.String.  The Identity (or Identities) for which to modify attributes
    can be passed via the command line.

    .OUTPUTS
    A PSObject with the requested attributes for the Identity.

    .EXAMPLE
    PS C:\> Set-ADAttribute -Identity wrightst -Attribute extensionAttribute2 -Value "hello" | Format-List

    Name                : wrightst
    DistinguishedName   : CN=wrightst,OU=Users,...
    extensionAttribute2 : hello

    The above example sets the "extensionAttribute2" attribute to the value "hello".
#>
################################################################################
function Set-ADAttribute {
    [CmdletBinding(SupportsShouldProcess=$true,
            ConfirmImpact="High")]
    param (
            [Parameter(Mandatory=$true,
                ValueFromPipeline=$true)]
            [ValidateNotNullOrEmpty()]
            [string]
            # Specifies the object to modify.
            $Identity,

            [Parameter(Mandatory=$true)]
            [ValidateNotNullOrEmpty()]
            [string]
            # Specifies which attribute to modify.
            $Attribute,

            [Parameter(Mandatory=$true)]
            [AllowNull()]
            [object[]]
            # Specifies the value to assign to the attribute.
            $Value,

            [Parameter(Mandatory=$false)]
            [string]
            # The domain controller to target.
            $DomainController
        )

    BEGIN {
        if ([String]::IsNullOrEmpty($DomainController)) {
            $pc = New-Object System.DirectoryServices.AccountManagement.PrincipalContext Domain
        } else {
            $pc = New-Object System.DirectoryServices.AccountManagement.PrincipalContext Domain, $DomainController
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
            Write-Error "Cannot find object $Identity in the current Active Directory domain."
            return
        } else {
            $props.DistinguishedName = $objUser.DistinguishedName
            Write-Verbose "Found DN: $($objUser.DistinguishedName)"
        }

        $dirEntry = $objUser.GetUnderlyingObject()

        $Error.Clear()
        if ($Value -eq $null) {
            $desc = "Clear attribute $Attribute for object $Identity"
            $caption = $desc
            $warning = "Are you sure you want to perform this action?`n"
            if ($PSCmdlet.ShouldProcess($desc, $warning, $caption)) {
                $dirEntry.PSBase.Properties.Item($Attribute).Clear()
                $dirEntry.PSBase.CommitChanges()
                if ($? -eq $true) {
                    Write-Verbose "Cleared $Attribute for $Identity"
                }
            }
        } else {
            try {
                $dirEntry.Properties.Item($Attribute).Clear()
                foreach ($val in $Value) {
                    $null = $dirEntry.Properties.Item($Attribute).Add($val)
                }
                $dirEntry.CommitChanges()
                if ($? -eq $true) {
                    Write-Verbose "Set $Attribute to $Value for $Identity"
                }
            } catch {
                Write-Error $_
            }
        }

        try {
            $dirEntry.RefreshCache()
            $props | Add-Member NoteProperty $Attribute $dirEntry.InvokeGet($Attribute)
        } catch {
            Write-Error "Attribute $Attribute does not exist for this object."
            return
        }

        $props
    }
}
