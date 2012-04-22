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
            # A string describing the category of this package.
            $Category,

            [System.Nullable[DateTime]]
            # When the package was created".
            $Created,

            [string]
            # The creator of the package.  This could be a user name, email address, etc.
            $Creator,

            [string]
            # The description of the package.
            $Description,

            [Guid]
            # The globally-unique identifier (GUID) of the package.
            $Identifier,

            [string[]]
            # An array of keywords describing the package.
            $Keywords,

            [System.Globalization.CultureInfo]
            # The language of the package.
            $Language,

            [string]
            # The user/email address/etc. of the person who last modified the package.
            $LastModifiedBy,

            [System.Nullable[DateTime]]
            # When the package was last modified.
            $Modified,

            [System.Nullable[int]]
            # The revision number of the package.
            $Revision,

            [string]
            # The package's subject matter.
            $Subject,

            [string]
            # The title of the package.
            $Title,

            [Version]
            # The version number of the package.
            $Version
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
                $Package = Get-Package $packagePath
            }
        } catch {
            throw $_
        }

        $BoundParameters = $PSCmdlet.MyInvocation.BoundParameters
        if ($BoundParameters.ContainsKey("Category")) {
            Write-Verbose "Setting Category to `"$Category`""
            $package.Category = $Category
        }
        if ($BoundParameters.ContainsKey("Created")) {
            Write-Verbose "Setting Created to $Created"
            $package.Created = $Created
        }
        if ($BoundParameters.ContainsKey("Creator")) {
            Write-Verbose "Setting Creator to `"$Creator`""
            $package.Creator = $Creator
        }
        if ($BoundParameters.ContainsKey("Description")) {
            Write-Verbose "Setting Description to `"$Description`""
            $package.Description = $Description
        }
        if ($Identifier -ne $null) {
            Write-Verbose "Setting Identifier to $Identifier"
            $package.Identifier = $Identifier
        }
        if ($BoundParameters.ContainsKey("Keywords")) {
            Write-Verbose "Setting Keywords to $Keywords"
            $package.Keywords = $Keywords -join ", "
        }
        if ($BoundParameters.ContainsKey("Language")) {
            Write-Verbose "Setting Language to $Language"
            $package.Language = $Language
        }
        if ($BoundParameters.ContainsKey("LastModifiedBy")) {
            Write-Verbose "Setting LastModifiedBy to $LastModifiedBy"
            $package.LastModifiedBy = $LastModifiedBy
        }
        if ($BoundParameters.ContainsKey("Modified")) {
            Write-Verbose "Setting Modified to $Modified"
            $package.Modified = $Modified
        }
        if ($BoundParameters.ContainsKey("Revision")) {
            Write-Verbose "Setting Revision to $Revision"
            $package.Revision = $Revision
        }
        if ($BoundParameters.ContainsKey("Subject")) {
            Write-Verbose "Setting Subject to $Subject"
            $package.Title = $Subject
        }
        if ($BoundParameters.ContainsKey("Title")) {
            Write-Verbose "Setting Title to $Title"
            $package.Title = $Title
        }
        if ($Version -ne $null) {
            Write-Verbose "Setting Version to $Version"
            $package.Version = $Version
        }

        $package.Flush()
    }
}
