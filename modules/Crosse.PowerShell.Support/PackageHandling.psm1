function Out-PackageFile {
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
            $BufferSize = 1 * 1024 * 1024,

            [Parameter(Mandatory=$false,
                ValueFromPipeline=$true)]
            [System.IO.FileSystemInfo[]]
            $InputObject,

            [switch]
            $ShowProgress
          )

    BEGIN {
        Write-Verbose "Performing initialization actions"
        $assembly = Get-ChildItem -Path 'C:\Program Files\Reference Assemblies\Microsoft\Framework' -Filter "WindowsBase.dll" -Recurse
        if ($assembly -eq $null) {
            throw New-Object System.IO.FileNotFoundException "Cannot find WindowsBase.dll"
        }
        Add-Type -Path $assembly.FullName

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
            if ($package -ne $null) {
                $package.Close()
            }
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
                    Write-Warning "Operation interrupted."
                    if ($package -ne $null) {
                        $package.Close()
                    }
                    return
                }
            }
        }
    }

    END {
        Write-Verbose "Cleaning up"
        if ($package -ne $null) {
            $package.Close()
        }
        if ($ShowProgress) {
            Write-Progress "Creating package $FilePath" -Activity "Done." -Completed
        }
    }
}
