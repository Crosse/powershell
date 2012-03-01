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

    .DESCRIPTION

    .INPUTS

    .OUTPUTS

    .EXAMPLE

    .EXAMPLE
#>
function New-Package {
    [CmdletBinding()]
    param (
            [Parameter(Mandatory=$true)]
            [string]
            $Name,

            [switch]
            $Force
          )
    try {
        if ((Split-Path -IsAbsolute $Name) -eq $true) {
            $packagePath = $Name
        } else {
            $packagePath = Join-Path (Get-Location -PSProvider "FileSystem") $Name
        }

        if ((Test-Path $packagePath) -and $Force -eq $false) {
            throw "File $packagePath already exists."
        }

        $package = [System.IO.Packaging.Package]::Open($packagePath, "Create")
    } catch {
        Close-Package $package
        throw $_
    }
    return $package
}
