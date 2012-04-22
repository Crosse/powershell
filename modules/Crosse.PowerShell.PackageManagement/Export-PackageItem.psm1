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
    Extracts an item from a package.

    .DESCRIPTION
    Extracts an item from a package.

    .INPUTS
    Export-PackageItem accepts either a System.IO.Packaging.Package object
    or the name of a package file.

    .OUTPUTS
    None.

    .EXAMPLE

    .LINK

#Requires -Version 2.0
#>

function Export-PackageItem {
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
                Position=0,
                ValueFromPipeline=$true,
                ParameterSetName="PackagePart")]
            [ValidateNotNull()]
            [System.IO.Packaging.PackagePart]
            # A PackagePart object.
            $PackagePart,

            [Parameter(Mandatory=$true,
                ValueFromPipelineByPropertyName=$true)]
            [ValidateNotNullOrEmpty()]
            [Alias("Uri")]
            [Uri]
            # The item to export from the package.
            $Path,

            [Parameter(Mandatory=$false)]
            [string]
            # The optional destination for the item.
            $Destination,

            [Parameter(Mandatory=$false)]
            [ValidateScript({ Test-Path $_ -PathType Container })]
            [System.IO.DirectoryInfo]
            $BasePath,

            [int]
            $BufferSize = 1MB
          )

    PROCESS {
        switch ($PSCmdlet.ParameterSetName) {
            "File" {
                $Package = Open-Package $PackagePath
                $PackagePart = Get-PackageItem -Name $Package | 
                    Where-Object {
                        [Uri]::UnescapeDataString($_.Uri) -eq $Path -or
                        $_.Uri -eq $Path
                    }
            }
            "Package" {
                $PackagePart = Get-PackageItem -Package $Package | 
                    Where-Object {
                        [Uri]::UnescapeDataString($_.Uri) -eq $Path -or
                        $_.Uri -eq $Path
                    }
            }
        }

        $normalizedPath = [Uri]::UnescapeDataString($Path)
        if ([String]::IsNullOrEmpty($BasePath)) {
            $base = $PWD
        } else {
            $base = $BasePath
        }

        if ([String]::IsNullOrEmpty($Destination)) {
            $dest = Join-Path $base $normalizedPath
        } elseif (Split-Path -IsAbsolute $Destination) {
            if ($BasePath) {
                throw "Cannot specify -BasePath with an explicit Destination."
            }
            $dest = $Destination
        } else {
            $dest = Join-Path $base $Destination
        }

        Write-Verbose "Exporting $normalizedPath as $dest"
        $parentPath = Split-Path -Parent $dest
        if (!(Test-Path $parentPath)) {
            Write-Verbose "Creating non-existent directory $parentPath"
            New-Item -ItemType Directory $parentPath
        }

        $packStream = $PackagePart.GetStream("Open", "Read")
        $len = $packStream.Length
        if ($len -lt $BufferSize) {
            $buffLength = $packStream.Length
        } else {
            $buffLength = $BufferSize
        }
        [byte[]]$buff = New-Object Byte[] $buffLength
        try {
            $stream = New-Object System.IO.FileStream($dest, "Create")
            $writer = New-Object System.IO.BinaryWriter($stream)

            while (($bytesRead = $packStream.Read($buff, 0, $buffLength)) -gt 0) {
                $writer.Write($buff, 0, $bytesRead)
            }
            $writer.Flush()
            Get-Item $dest
        } catch {
            throw
        } finally {
            if ($packStream) {
                $packStream.Close()
                $packStream.Dispose()
            }
            if ($writer) {
                $writer.Close()
                $writer.Dispose()
            }
        }
    }
}
