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
    Gets the properties of a package file.

    .DESCRIPTION
    Gets the properties of a package file.

    .INPUTS
    Get-PackageProperty accepts either a System.IO.Packaging.Package object
    or the name of a package file.

    .OUTPUTS
    The properties of the package file.

    .EXAMPLE
    Get-PackageProperty $pack


    Creator        : wrightst
    Title          : test.zip
    Subject        :
    Category       :
    Keywords       :
    Description    :
    ContentType    :
    ContentStatus  :
    Version        : 1.0
    Revision       :
    Created        : 3/1/2012 10:39:39 PM
    Modified       : 3/1/2012 10:39:39 PM
    LastModifiedBy : wrightst
    LastPrinted    :
    Language       :
    Identifier     :

    .LINK

#Requires -Version 2.0
#>

function Get-PackageProperty {
    [CmdletBinding()]
    param (
            [Parameter(Mandatory=$true,
                ValueFromPipeline=$true,
                Position=0,
                ParameterSetName="File")]
            [string]
            # The path to a package file.
            $Name,

            [Parameter(Mandatory=$true,
                ValueFromPipeline=$true,
                Position=0,
                ParameterSetName="Package")]
            [Crosse.PowerShell.PackageManagement.PackageFile]
            # A PackageFile object.
            $Package
          )

    PROCESS {
        try {
            if ([String]::IsNullOrEmpty($Name) -eq $false) {
                if ((Split-Path -IsAbsolute $Name) -eq $true) {
                    $packagePath = Resolve-Path $Name -ErrorAction Stop
                } else {
                    $packagePath = Resolve-Path `
                        (Join-Path (Get-Location -PSProvider "FileSystem") $Name) -ErrorAction Stop
                }
                $Package = Open-Package $packagePath
            }

            $Package.Package.PackageProperties
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
