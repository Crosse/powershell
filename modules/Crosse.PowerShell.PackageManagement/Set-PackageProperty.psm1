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
    Sets various properties of a package.

    .DESCRIPTION
    Sets various properties of a package.

    .INPUTS
    Set-PackageProperties accepts either a System.IO.Packaging.Package object
    or the name of a package file.

    .OUTPUTS
    None. Set-PackageProperties does not return anything.

    .EXAMPLE
    Set-PackageProperty $pack -Description "Description!"

    This example illustrates setting the Description property of the package
    stored in $pack to "Description!"

    .EXAMPLE
    Set-PackageProperty .\test.zip -Modified (Get-Date) -Verbose
    VERBOSE: Setting Modified to 03/01/2012 23:25:12


    This example illustrates setting the Modified property of the package to the
    current time and date.

    .LINK

#Requires -Version 2.0
#>

function Set-PackageProperty {
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
            [ValidateNotNull()]
            [Crosse.PowerShell.PackageManagement.PackageFile]
            # A PackageFile object.
            $Package,

            [string]
            # The creator of the package.  This could be a user name, email address, etc.
            $Creator,

            [string]
            # The title of the package.
            $Title,

            [string]
            # The description of the package.
            $Description,

            [string]
            # The version number of the package.
            $Version,

            [string]
            # The revision number of the package.
            $Revision,

            [DateTime]
            # When the package was created.
            $Created,

            [DateTime]
            # When the package was last modified.
            $Modified,

            [string]
            # The user/email address/etc. of the person who last modified the package.
            $LastModifiedBy,

            [Guid]
            # The globally-unique identifier (GUID) of the package.
            $Identifier
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

            if ($package.Package.PackageProperties -ne $null) {
                if ([String]::IsNullOrEmpty($Creator) -eq $false) {
                    Write-Verbose "Setting Creator to $Creator"
                    $package.Package.PackageProperties.Creator = $Creator
                }
                if ([String]::IsNullOrEmpty($Title) -eq $false) {
                    Write-Verbose "Setting Title to $Title"
                    $package.Package.PackageProperties.Title = $Title
                }
                if ([String]::IsNullOrEmpty($Description) -eq $false) {
                    Write-Verbose "Setting Description to $Description"
                    $package.Package.PackageProperties.Description = $Description
                }
                if ([String]::IsNullOrEmpty($Version) -eq $false) {
                    Write-Verbose "Setting Version to $Version"
                    $package.Package.PackageProperties.Version = $Version
                }
                if ([String]::IsNullOrEmpty($Revision) -eq $false) {
                    Write-Verbose "Setting Revision to $Revision"
                    $package.Package.PackageProperties.Revision = $Revision
                }
                if ($Created) {
                    Write-Verbose "Setting Created to $Created"
                    $package.Package.PackageProperties.Created = $Created
                }
                if ($Modified) {
                    Write-Verbose "Setting Modified to $Modified"
                    $package.Package.PackageProperties.Modified = $Modified
                }
                if ([String]::IsNullOrEmpty($LastModifiedBy) -eq $false) {
                    Write-Verbose "Setting LastModifiedBy to $LastModifiedBy"
                    $package.Package.PackageProperties.LastModifiedBy = $LastModifiedBy
                }
                if ([String]::IsNullOrEmpty($Identifier) -eq $false) {
                    Write-Verbose "Setting Identifier to $Identifier"
                    $package.Package.PackageProperties.Identifier = $Identifier
                }

                $package.Package.Flush()
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
