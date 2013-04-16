################################################################################
#
# Copyright (c) 2013 Seth Wright <seth@crosse.org>
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

################################################################################
<#
    .SYNOPSIS
    Transcodes a source video file.

    .DESCRIPTION
    Transcodes a source video file into a destination MP4 or MKV file.

    .INPUTS
    System.String. The names of multiple source files to transcode can be passed via the command line.

    .OUTPUTS
    None.

#>
################################################################################
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

            [Parameter(Mandatory=$false,
                ParameterSetName="Bluray")]
            [ValidateSet("480p", "720p", "1080p")]
            [string]
            $VideoPreset,

            [Parameter(Mandatory=$true,
                ParameterSetName="ScanOnly")]
            [switch]
            $ScanOnly

            [Parameter(Mandatory=$false)]
            [string]
            # The path to HandbrakeCLI.exe.  By default this is "C:\Program Files\Handbrake\HandbrakeCLI.exe".
            $HandbrakeCLIPath = "C:\Program Files\Handbrake\HandbrakeCLI.exe",

            [Parameter(Mandatory=$false)]
            [string]
            # The path to the MediaInfo.exe CLI program (NOT the GUI application, which unfortunately has the same name).  The default is to use the binaries in the same directory as this cmdlet's module manifest.
            $MediaInfoCLIPath = (Join-Path $PSScriptRoot "MediaInfo.exe"),
          )

    begin {
        Resolve-Path $HandbrakeCLIPath -ErrorAction Stop | Out-Null
        Write-Verbose "Found HandbrakeCLI: $HandbrakeCLIPath"

        Resolve-Path $MediaInfoCLIPath -ErrorAction Stop | Out-Null
        Write-Verbose "Found MediaInfo: $MediaInfoCLIPath"

        $outPath = Resolve-Path $OutputPath -ErrorAction Stop

        $handbrakeOptions = @(
                # Set output format
                '--format mp4',
                # Add chapter markers
                '--markers',
                # Use 64-bit mp4 files that can hold more than 4GB.
                '--large-file',
                # Optimize mp4 files for HTTP streaming
                '--optimize',
                # Set video library encoder
                '--encoder x264',
                # advanced encoder options in the same style as mencoder
                '--encopts "b-adapt=2"',
                # Set video quality
                '--quality 18',
                # Set video framerate
                '--rate 30',
                # Select peak-limited frame rate control.
                '--pfr',
                # Set audio codec to use when it is not possible to copy an
                # audio track without re-encoding.
                '--audio-fallback ffac3',
                # Store pixel aspect ratio with specified width
                '--loose-anamorphic',
                # Set the number you want the scaled pixel dimensions to divide
                # cleanly by.
                '--modulus 2',
                # Selectively deinterlaces when it detects combing
                '--decomb',
                # Select subtitle track(s), separated by commas.  A special
                # track name "scan" adds an extra 1st pass.  This extra pass
                # scans subtitles matching the language of the first audio or
                # the language selected by --native-language.  The one that's
                # only used 10 percent of the time or less is selected. This
                # should locate subtitles for short foreign language segments.
                # Best used in conjunction with --subtitle-forced.
                '--subtitle scan',
                # Specifiy your language preference.
                '--native-language eng'
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

            $command = "& '$MediaInfoCLIPath' --Output=XML `"$($inputFile.FullName)`""
            [xml]$info = Invoke-Expression "$command"

            $audio = $info.MediaInfo.File.Track | ? { $_.type -eq "Audio" }
            if ($audio -eq $null) {
                Write-Error "Error getting audio track information from source."
                return
            }

            $mainAudio = $audio | ? { $_.Default -eq "Yes" }
            switch ($mainAudio.Format) {
                "AC-3" {
                    $audioTitle = $mainAudio.Title
                    $audioOptions = @(
                            '--aencoder "copy:ac3,copy:aac"',
                            '-A "Dolby Digital $audioTitle,Dolby Pro Logic II"'
                            )
                }
                "DTS" {
                    if ($mainAudio.Format_profile -match '^MA') {
                        $audioOptions = @(
                                '--aencoder "copy:ac3,copy:dtshd,copy:aac"',
                                '-A "Dolby Digital 5.1,DTS-HD MA,Dolby Pro Logic II"'
                                )
                    } else {
                        $audioOptions = @(
                                '--aencoder "copy:dts,copy:ac3,copy:aac"',
                                '-A "DTS","Dolby Digital 5.1,Dolby Pro Logic II"'
                                )
                    }
                }
            }

            if ([String]::IsNullOrEmpty($VideoPreset)) {
                $video = $info.MediaInfo.File.Track | ? { $_.type -eq "Video" }
                if ($audio -eq $null) {
                    Write-Error "Error getting audio track information from source."
                    return
                }
            }
            else {
                switch ($VideoPreset) {
                    '480p' { $videoOptions = '--maxWidth 480' }
                    '720p' { $videoOptions = '--maxWidth 1280' }
                    '1080p' { $videoOptions = '--maxWidth 1920' }
                }
            }

            if ($ScanOnly) {
                $fileOptions = "--scan --input `"$($inputFile.FullName)`""
            } else {
                $fileOptions = "--input `"$($inputFile.FullName)`" --output `"$outputFile`""
            }
            $command = "& '$HandbrakeCLIPath' $fileOptions "
            $command += $handbrakeOptions -join " "
            $command += " "
            $command += $audioOptions -join " "
            $command += " "
            $command += $videoOptions -join " "
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
