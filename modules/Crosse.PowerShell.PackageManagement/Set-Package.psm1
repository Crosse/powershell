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

function Set-Package {
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

            if ($Creator -ne $null) {
                Write-Verbose "Setting Creator to $Creator"
                $package.Creator = $Creator
            }
            if ($Title -ne $null) {
                Write-Verbose "Setting Title to $Title"
                $package.Title = $Title
            }
            if ($Description -ne $null) {
                Write-Verbose "Setting Description to $Description"
                $package.Description = $Description
            }
            if ($Version -ne $null) {
                Write-Verbose "Setting Version to $Version"
                $package.Version = $Version
            }
            if ($Revision -ne $null) {
                Write-Verbose "Setting Revision to $Revision"
                $package.Revision = $Revision
            }
            if ($Created) {
                Write-Verbose "Setting Created to $Created"
                $package.Created = $Created
            }
            if ($Modified) {
                Write-Verbose "Setting Modified to $Modified"
                $package.Modified = $Modified
            }
            if ($LastModifiedBy -ne $null) {
                Write-Verbose "Setting LastModifiedBy to $LastModifiedBy"
                $package.LastModifiedBy = $LastModifiedBy
            }
            if ($Identifier -ne $null) {
                Write-Verbose "Setting Identifier to $Identifier"
                $package.Identifier = $Identifier
            }

            $package.Flush()
        } catch {
            throw $_
        }
    }
}