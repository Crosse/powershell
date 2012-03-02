################################################################################
#
# Copyright (c) 2012 Seth Wright <wrightst@jmu.edu>
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

<#
    .SYNOPSIS
    Removes an item from a package.

    .DESCRIPTION
    Removes an item from a package.

    .INPUTS
    Remove-PackageItem accepts either a System.IO.Packaging.Package object
    or the name of a package file.

    .OUTPUTS
    None.

    .EXAMPLE

    .LINK

#Requires -Version 2.0
#>

function Remove-PackageItem {
    [CmdletBinding()]
    param (
            [Parameter(Mandatory=$true,
                Position=0,
                ParameterSetName="File")]
            [string]
            # The path to a package file.
            $Name,

            [Parameter(Mandatory=$true,
                Position=0,
                ParameterSetName="Package")]
            [ValidateNotNull()]
            [Crosse.PowerShell.PackageManagement.PackageFile]
            # A PackageFile object.
            $Package,

            [Parameter(Mandatory=$true,
                Position=0,
                ValueFromPipeline=$true,
                ParameterSetName="PackagePart")]
            [ValidateNotNull()]
            [System.IO.Packaging.PackagePart]
            # A PackagePart object.
            $PackagePart,

            [Parameter(Mandatory=$true,
                ValueFromPipelineByPropertyName=$true)]
            [ValidateNotNullOrEmpty()]
            [Alias("Uri")]
            [Uri]
            # The item to remove from the package.
            $Path
          )

    PROCESS {
        try {
            switch ($PSCmdlet.ParameterSetName) {
                "File" {
                    $Package = Open-Package $Name
                    $pack = $Package.Package
                    break;
                }
                "Package" {
                    $pack = $Package.Package
                    break;
                }
                "PackagePart" {
                    $pack = $PackagePart.Package
                    break;
                }
            }

            if ($pack.PartExists($Path)) {
                Write-Verbose "Deleting $Path from package"
                $pack.DeletePart($Path)
            } else {
                Write-Warning "Item $ItemPath does not exist in package."
            }
        } catch {
            if (![String]::IsNullOrEmpty($Name) -and $Package -ne $null) {
                Close-Package $Package
            }
            throw $_
        }
        finally {
            if (![String]::IsNullOrEmpty($Name)) {
                Close-Package $Package
            }
        }
    }
}
