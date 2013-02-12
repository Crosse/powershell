function Out-M4V {
    [CmdletBinding()]
    param (
            [Parameter(Mandatory=$true,
                ValueFromPipeline=$true)]
            [System.IO.FileInfo]
            $File,

            [Parameter(Mandatory=$false)]
            [System.IO.DirectoryInfo]
            $OutputPath = (Get-Location).Path
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
        try {
            $inputFile = [System.IO.FileInfo](Resolve-Path $File -ErrorAction Stop).Path
            $outputFile = Join-Path $outPath $inputFile.Name.Replace($inputFile.Extension, ".m4v")

            $fileOptions = "--input `"$inputFile`" --output `"$outputFile`""
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
