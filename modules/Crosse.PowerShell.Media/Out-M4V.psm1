function Out-M4V {
    [CmdletBinding()]
    param (
            [Parameter(Mandatory=$true,
                ValueFromPipeline=$true)]
            [object]
            $File,

            [Parameter(Mandatory=$false)]
            [System.IO.DirectoryInfo]
            $OutputPath = (Get-Location).Path,

            [Parameter(Mandatory=$true,
                ParameterSetName="ScanOnly")]
            [switch]
            $ScanOnly
          )

    begin {
        $handbrake = "C:\Program Files\Handbrake\HandbrakeCLI.exe"
        Resolve-Path $handbrake -ErrorAction Stop | Out-Null
        $outPath = Resolve-Path $OutputPath -ErrorAction Stop

        $handbrakeOptions = @(
                '--format mp4',
                '--markers',
                '--large-file',
                '--optimize',
                '--encoder x264',
                '--encopts "b-adapt=2"',
                '--quality 20',
                '--rate 30',
                '--pfr',
                '--aencoder "ffaac,copy:ac3"',
                '--audio-copy-mask none',
                '--audio-fallback ffac3',
                '--ab "160,0"',
                '--mixdown "dpl2,auto"',
                '--arate "Auto,Auto"',
                '--drc "0,0"',
                '--gain "0,0"',
                '--aname English',
                '--width 720',
                '--loose-anamorphic',
                '--modulus 2',
                '--decomb',
                '--subtitle scan'
                )

        if ($Verbose) {
            $handbrakeOptions += "--verbose"
        } else {
            $handbrakeOptions += "--verbose 0"
        }

    }
    process {
        Write-Verbose "Working on $File"
        try {
            if ($File -is [System.IO.FileInfo]) {
                $inputFile = $File
            } else {
                $inter = Resolve-Path $File
                $inputFile = [System.IO.FileInfo]$inter.Path
            }
            Write-Verbose "Input File: $inputFile"
            $outputFile = Join-Path $outPath $inputFile.Name.Replace($inputFile.Extension, ".m4v")

            $fileOptions = "--input `"$inputFile`" --output `"$outputFile`""
            if ($ScanOnly) {
                $fileOptions = "--scan --input `"$($inputFile.FullName)`"" 
            } else {
                $fileOptions = "--input `"$($inputFile.FullName)`" --output `"$outputFile`""
            }
            $command = "& '$handbrake' $fileOptions "
            $command += $handbrakeOptions -join " "
            Write-Verbose $command
            Invoke-Expression "$command"
        }
        catch
        {
            throw
        }
    }
    end {
    }
}
