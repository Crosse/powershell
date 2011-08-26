################################################################################
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


function Set-ADAttribute {
    [CmdletBinding(SupportsShouldProcess=$true,
            ConfirmImpact="High")]
    param (
            [Parameter(Mandatory=$true,
                ValueFromPipeline=$true)]
            [ValidateNotNullOrEmpty()]
            [string]
            # Specifies the object should be modified.
            $Identity,

            [Parameter(Mandatory=$true)]
            [string]
            # Specifies which attribute to modify.
            $Attribute,

            #[Parameter(Mandatory=$true)]
            [string]
            # Specifies the value to assign to the attribute.
            $Value
        )

    BEGIN {
        $assembly = [Reflection.Assembly]::LoadWithPartialName("System.DirectoryServices.AccountManagement")
        if ($assembly -eq $null) {
            Write-Error "Could not load the System.DirectoryServices.AccountManagement assembly"
            return
        }

        $pc = New-Object System.DirectoryServices.AccountManagement.PrincipalContext Domain
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
        $dirEntry.InvokeSet($Attribute, $Value)
        $dirEntry.PSBase.CommitChanges()
        if ($? -eq $true) {
            Write-Verbose "Set $Attribute to $Value"
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
