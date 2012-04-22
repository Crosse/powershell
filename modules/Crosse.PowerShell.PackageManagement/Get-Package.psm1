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
    Gets a package file.

    .DESCRIPTION
    Gets a package file.

    .INPUTS
    System.String
        You can pipe a string that contains a path to Get-Package.

    .OUTPUTS
    Get-Package returns a Crosse.PowerShell.PackageManagement.PackageFile
    object referencing the package.

    .EXAMPLE
    Get-Package test.zip

    This example illustrates getting a package.

    .LINK

#Requires -Version 2.0
#>
function Get-Package {
    [CmdletBinding()]
    param (
            [Parameter(Mandatory=$true,
                ValueFromPipeline=$true)]
            [string]
            # The path to a package file.
            $PackagePath
          )

    if ((Split-Path -IsAbsolute $PackagePath) -eq $true) {
        $packagePath = Resolve-Path $PackagePath -ErrorAction Stop
    } else {
        $packagePath = Resolve-Path `
            (Join-Path (Get-Location -PSProvider "FileSystem") $PackagePath) -ErrorAction Stop
    }

    New-Object Crosse.PowerShell.PackageManagement.PackageFile($packagePath, "Open")
}
