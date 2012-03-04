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
    Closes a package file.

    .DESCRIPTION
    Closes a package file.

    .INPUTS
    A System.IO.Packaging.Package object.

    .OUTPUTS
    None.

    .EXAMPLE
    Close-Package $package

    This example illustrates closing a package that had previously been
    opened and saved into the variable "$package".

    .LINK 

#Requires -Version 2.0 
#>
    
function Close-Package {
    [CmdletBinding()]
    param (
            [Parameter(Mandatory=$true,
                Position=0,
                ValueFromPipeline=$true,
                ParameterSetName="File")]
            [string]
            # The path to a package file.
            $Name,

            [Parameter(Mandatory=$true,
                Position=0,
                ValueFromPipeline=$true,
                ParameterSetName="Package")]
            [AllowNull()]
            [Crosse.PowerShell.PackageManagement.PackageFile]
            # A PackageFile object.
            $Package
          )

    switch ($PSCmdlet.ParameterSetName) {
        "File" {
            [GC]::Collect()
        }
        "Package" {
            if ($Package -ne $null -and $Package.Package -ne $null) {
                $Package.Close()
            }
        }
    }
}
