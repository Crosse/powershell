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
    Close-Package $package

    This example illustrates closing a package that had previously been
    opened and saved into the variable "$package".

    .LINK

#Requires -Version 2.0
#>

function Out-Package {
    [CmdletBinding()]
    param (
            [Parameter(Mandatory=$true)]
            [ValidateNotNullOrEmpty()]
            [string]
            $FilePath,

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

            if ( (Split-Path -IsAbsolute $FilePath) -eq $true) {
                $packagePath = $FilePath
            } else {
                $packagePath = Join-Path (Get-Location -PSProvider "FileSystem") $FilePath
            }

            $package = [System.IO.Packaging.Package]::Open($packagePath, $fileMode)
            $basePath = $null
        } catch {
            Close-Package $package
            throw $_
        }

        $totalObjects = @($InputObject).Count
        $totalProgress = 0

        if ($ShowProgress) {
            Write-Progress "Compressing" -Activity "Creating package $FilePath"
        }
    }

    PROCESS {
        $i = 0
        foreach ($fso in $InputObject) {
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
                if ($package.PartExists($uri)) {
                    Write-Verbose "$uri already exists in package; deleting"
                    $package.DeletePart($uri)
                }
                $part = $package.CreatePart($uri, "", "Normal")
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
                        Write-Progress "$percent% Complete" -Activity "Creating package $FilePath" -PercentComplete $percent
                    } else {
                        Write-Progress "Compressed $relPath" -Activity "Creating package $FilePath"
                    }
                }
                $i++
                $interrupted = $false
            } catch {
                Write-Error $_
                if ($package -ne $null) {
                    $package.Close()
                }
                return
            } finally {
                if ($srcStream -ne $null) {
                    $srcStream.Close()
                }
                if ($destStream -ne $null) {
                    $destStream.Close()
                }
                if ($interrupted) {
                    Close-Package $package
                    throw "Operation interrupted."
                }
            }
        }
    }

    END {
        Write-Verbose "Cleaning up"
        Close-Package $package
        if ($ShowProgress) {
            Write-Progress "Creating package $FilePath" -Activity "Done." -Completed
        }
    }
}
