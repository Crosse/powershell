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
    Adds an item to a package.

    .DESCRIPTION
    Adds an item to a package, optionally with a specified path in the package.

    .INPUTS
    Add-PackageItem accepts either a System.IO.Packaging.Package object
    or the name of a package file.

    .OUTPUTS
    None.

    .EXAMPLE
    Add-PackageItem -Package $pack -Source .\Add-PackageItem.psm1 -AllowUpdate

    This example illustrates how to add a file to a package if the source already
    exists in the package.

    .LINK

#Requires -Version 2.0
#>

function Add-PackageItem {
    [CmdletBinding()]
    param (
            [Parameter(Mandatory=$true,
                Position=0,
                ParameterSetName="File")]
            [string]
            # The path to a package file.
            $PackagePath,

            [Parameter(Mandatory=$true,
                Position=0,
                ParameterSetName="Package")]
            [ValidateNotNull()]
            [Crosse.PowerShell.PackageManagement.PackageFile]
            # A PackageFile object.
            $Package,

            [Parameter(Mandatory=$true,
                ValueFromPipeline=$true,
                Position=1)]
            [ValidateScript({Test-Path $_.FullName -PathType Leaf})]
            [System.IO.FileInfo]
            # The file to add to the package.
            $Source,

            [Parameter(Mandatory=$false,
                    Position=2)]
            [Uri]
            # The Uri of the file to use when adding it to the package.
            $Destination,

            [switch]
            # If the file already exists in the package, allow it to be updated.
            $AllowUpdate,

            [Parameter(Mandatory=$false)]
            [ValidateSet("NotCompressed", "Normal", "Maximum", "Fast", "SuperFast")]
            [string]
            # The type of compression to use when adding the item to the package.
            $CompressionOption = "Normal",

            [Parameter(Mandatory=$false)]
            [int]
            # The buffer size to use.
            $BufferSize = 1MB
          )

    PROCESS {
        try {
            if ([String]::IsNullOrEmpty($PackagePath) -eq $false) {
                $Package = Get-Package $PackagePath
            }

            if ($Destination) {
                $destPath = $Destination
            } else {
                $destPath = $Source.Name
            }

            $pack = $Package.Package

            Write-Verbose "Adding $Source to package as $destPath"
            $uri = [System.IO.Packaging.PackUriHelper]::CreatePartUri($destPath)
            if ($pack.PartExists($uri)) {
                if ($AllowUpdate) {
                    Write-Verbose "$uri already exists in package; deleting"
                    $pack.DeletePart($uri)
                } else {
                    throw "$uri already exists in package and -AllowUpdate not specified"
                }
            }
            $part = $pack.CreatePart($uri, "", "Normal")
            $srcStream = New-Object System.IO.FileStream($Source.FullName, "Open", "Read")
            $destStream = $part.GetStream()

            $len = $srcStream.Length
            if ($len -lt $BufferSize) {
                $buffLength = $srcStream.Length
            } else {
                $buffLength = $BufferSize
            }
            [byte[]]$buff = New-Object Byte[] $buffLength

            $totalBytesRead = 0
            while ( ($bytesRead = $srcStream.Read($buff, 0, $buffLength)) -gt 0) {
                $destStream.Write($buff, $offset, $bytesRead)
                $totalBytesRead += $bytesRead
            }
            $destStream.Flush()
        } catch {
            if ($uri -and $pack.PartExists($uri)) {
                $pack.DeletePart($uri)
            }
            if (![String]::IsNullOrEmpty($PackagePath) -and $Package -ne $null) {
                Close-Package $Package
            }
            throw $_
        }
        finally {
            if ($srcStream -ne $null) {
                $srcStream.Close()
            }
            if ($destStream -ne $null) {
                $destStream.Close()
            }
            if (![String]::IsNullOrEmpty($PackagePath)) {
                Close-Package $Package
            }
        }
    }
}
