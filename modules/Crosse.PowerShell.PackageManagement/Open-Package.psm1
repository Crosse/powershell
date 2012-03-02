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
    Opens a package file.

    .DESCRIPTION
    Opens a package file and returns a reference to it.

    .INPUTS
    The path to the package file to open.

    .OUTPUTS
    A Crosse.PowerShell.PackageManagement.PackageFile pointing to the package.

    .EXAMPLE
    $pack = Open-Package .\test.zip

    This example illustrates opening a package file.

    .LINK

#Requires -Version 2.0
#>

function Open-Package {
    [CmdletBinding()]
    param (
            [Parameter(Mandatory=$true)]
            [string]
            # The path to the package file to open.
            $Name
          )

    PROCESS {
        try {
            if ((Split-Path -IsAbsolute $Name) -eq $true) {
                $packagePath = Resolve-Path $Name -ErrorAction Stop
            } else {
                $packagePath = Resolve-Path `
                    (Join-Path (Get-Location -PSProvider "FileSystem") $Name) -ErrorAction Stop
            }

            New-Object Crosse.PowerShell.PackageManagement.PackageFile($packagePath, "Open")
        } catch {
            Close-Package $package
            throw $_
        }
    }
}
