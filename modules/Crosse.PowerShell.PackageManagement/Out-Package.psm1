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
    Closes a package file.

    .DESCRIPTION
    Closes a package file.

    .INPUTS
    A System.IO.Packaging.Package object.

    .OUTPUTS
    None.

    .EXAMPLE
    Close-Package $pack

    This example illustrates closing a package that had previously been
    opened and saved into the variable "$pack".

    .LINK

#Requires -Version 2.0
#>

function Out-Package {
    [CmdletBinding()]
    param (
            [Parameter(Mandatory=$true)]
            [ValidateNotNullOrEmpty()]
            [string]
            $PackagePath,

            [switch]
            $AddOrUpdate,

            [Parameter(Mandatory=$false)]
            [ValidateSet("NotCompressed", "Normal", "Maximum", "Fast", "SuperFast")]
            [string]
            $CompressionOption = "Normal",

            [Parameter(Mandatory=$false)]
            [int]
            $BufferSize = 1MB,

            [Parameter(Mandatory=$false,
                ValueFromPipeline=$true)]
            [System.IO.FileSystemInfo[]]
            $InputObject,

            [switch]
            $ShowProgress
          )

    BEGIN {
        Write-Verbose "Performing initialization actions"
        try {
            if ($AddOrUpdate) {
                $fileMode = "OpenOrCreate"
            } else {
                $fileMode = "Create"
            }

            if ( (Split-Path -IsAbsolute $PackagePath) -eq $true) {
                $packPath = $PackagePath
            } else {
                $packPath = Join-Path (Get-Location -PSProvider "FileSystem") $PackagePath
            }

            if (Test-Path $packPath) {
                if ($AddOrUpdate) {
                    $package = Get-Package $packPath
                } else {
                    throw "Package exists, and -AddOrUpdate wasn't specified."
                }
            } else {
                $package = New-Package $packPath
            }
            $pack = $package.GetUnderlyingPackage()
            $basePath = $null
        } catch {
            $package.CloseUnderlyingPackage()
            throw
        }

        $totalObjects = @($InputObject).Count
        $totalProgress = 0

        if ($ShowProgress) {
            Write-Progress "Compressing" -Activity "Creating package $PackagePath"
        }
    }

    PROCESS {
        $i = 0
        foreach ($fso in $InputObject) {
            Write-Verbose $fso
            if ($fso -eq $null) {
                continue
            }
            if ($fso.GetType().Name -eq "FileInfo") {
                if ([String]::IsNullOrEmpty($basePath)) {
                    $basePath = $fso.DirectoryName
                }
            } elseif ($fso.GetType().Name -eq "DirectoryInfo") {
                if ([String]::IsNullOrEmpty($basePath)) {
                    $basePath = $fso.Parent.Parent.FullName
                }
                continue
            }

            $relPath = $fso.FullName.Replace($basePath, "")
            $relPath = "." + $relPath

            try {
                $interrupted = $true
                Write-Verbose "Adding $relPath"
                $uri = [System.IO.Packaging.PackUriHelper]::CreatePartUri($relPath)
                if ($pack.PartExists($uri)) {
                    Write-Verbose "$uri already exists in package; deleting"
                    $pack.DeletePart($uri)
                }
                $part = $pack.CreatePart($uri, "", "Normal")
                $srcStream = New-Object System.IO.FileStream($fso.FullName, "Open", "Read")
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
                    if ($ShowProgress -and
                            $totalObjects -gt 1 -and
                            $len -gt $buffLength) {
                        Write-Progress -Activity "Compressing $relPath" `
                                       -PercentComplete ($totalBytesRead/$len*100) `
                                       -Status "Compressing" -Id 1
                    }
                    $destStream.Write($buff, $offset, $bytesRead)
                    $totalBytesRead += $bytesRead
                }
                if ($ShowProgress) {
                    if ($totalObjects -gt 1) {
                        $percent = [Math]::Round($i/$totalObjects*100, 0)
                        Write-Progress "$percent% Complete" -Activity "Creating package $PackagePath" -PercentComplete $percent
                    } else {
                        Write-Progress "Compressed $relPath" -Activity "Creating package $PackagePath"
                    }
                }
                $i++
                $interrupted = $false
            } catch {
                $err = $_
                if ($pack -ne $null) {
                    $pack.Close()
                }
                throw $err
            } finally {
                if ($srcStream -ne $null) {
                    $srcStream.Close()
                }
                if ($destStream -ne $null) {
                    $destStream.Close()
                }
                if ($interrupted) {
                    $pack.CloseUnderlyingPackage()
                    throw "Operation interrupted."
                }
            }
        }
    }

    END {
        Write-Verbose "Cleaning up"
        $package.CloseUnderlyingPackage()
        if ($ShowProgress) {
            Write-Progress "Creating package $PackagePath" -Activity "Done." -Completed
        }
    }
}
