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
    Creates a new package file.

    .DESCRIPTION
    Creates a new package file.

    .INPUTS
    None.

    .OUTPUTS
    New-Package returns a Crosse.PowerShell.PackageManagement.PackageFile
    object referencing the package.

    .EXAMPLE
    New-Package test.zip

    This example illustrates creating a new package.

    .EXAMPLE
    New-Package test.zip -Force

    This example illustrates creating a new package by overwriting any previous
    package file with the same name.  This will still fail is the existing package
    is currently open.

    .LINK

#Requires -Version 2.0
#>
function New-Package {
    [CmdletBinding()]
    param (
            [Parameter(Mandatory=$true)]
            [ValidateNotNullOrEmpty()]
            [string]
            # The path to the package file to create.
            $PackagePath,

            [Guid]
            # A globally-unique identifier (GUID) used to identify the package.
            $Identifier = [Guid]::NewGuid(),

            [switch]
            # Specifies that if a package with the same name already exists,
            # it should be overwritten.
            $Force
          )

    try {
        if ((Split-Path -IsAbsolute $PackagePath) -eq $true) {
            $packagePath = Resolve-Path $PackagePath
        } else {
            $packagePath = Join-Path (Get-Location -PSProvider "FileSystem") $PackagePath
        }

        if ((Test-Path $packagePath) -and $Force -eq $false) {
            throw "File $packagePath already exists."
        }

        $package = New-Object Crosse.PowerShell.PackageManagement.PackageFile($packagePath, "Create")
        $creator = (Get-Item Env:\USERNAME).Value
        $now = Get-Date
        Set-Package -Package $package `
                    -Creator $creator `
                    -Title $PackagePath `
                    -Version "1.0" `
                    -Created $now `
                    -Modified $now `
                    -Identifier $Identifier `
                    -LastModifiedBy $creator
        Get-Package $packagePath
    } catch {
        throw
    }
}
