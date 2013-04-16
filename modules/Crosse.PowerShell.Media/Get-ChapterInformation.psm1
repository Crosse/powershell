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
    Retrieves video chapter information from online sources.

    .DESCRIPTION
    Retrieves video chapter information from online sources.

    .INPUTS
    None.

    .OUTPUTS
    None.

#>
################################################################################
function Get-ChapterInformation {
    [CmdletBinding()]
    param (
            [Parameter(Mandatory=$true)]
            [ValidateNotNullOrEmpty()]
            [string]
            # The title to search for.
            $Title,

            [Parameter(Mandatory=$false)]
            [int]
            # The number of chapters.
            $ChapterCount,

            [Parameter(Mandatory=$false)]
            [switch]
            # Indicates whether to return all results or just a single result deemed the "best".
            $BestResult = $true,

            [Parameter(Mandatory=$true,
                ParameterSetName="ChaptersDb")]
            [switch]
            $UseChaptersDb,

            [Parameter(Mandatory=$true,
                ParameterSetName="TagChimp")]
            [switch]
            $UseTagChimp,

            [Parameter(Mandatory=$true,
                ParameterSetName="UseBoth")]
            [switch]
            $UseBothServices,

            [Parameter(Mandatory=$true,
                ParameterSetName="ChaptersDb")]
            [Parameter(Mandatory=$true,
                ParameterSetName="UseBoth")]
            [string]
            # Your API Key for the ChaptersDb.org website.
            $ChaptersDbApiKey,

            [Parameter(Mandatory=$true,
                ParameterSetName="TagChimp")]
            [Parameter(Mandatory=$true,
                ParameterSetName="UseBoth")]
            [string]
            # Your API Key for the tagChimp.com website.
            $TagChimpApiKey
          )

    $escapedTitle = [Uri]::EscapeUriString($Title).Replace("*", "")

    $results = @()
    if ([String]::IsNullOrEmpty($ChaptersDbApiKey) -eq $false) {
        $chaptersDbApi = "http://chapterdb.org/chapters/search"
        $chaptersDbRequest = "{0}?title={1}&chapterCount={2}" -f $chaptersDbApi, $escapedTitle, $ChapterCount

        $response = Invoke-WebRequest -Uri $chaptersDbRequest -Method GET -Headers @{ ApiKey = $ChaptersDbApiKey }
        if ($response.StatusCode -ne 200) {
            Write-Error "Request unsuccessful. ($($response.StatusCode), $($response.StatusDescription))"
            return
        }

        $info = ([xml]($response.Content)).results.chapterInfo

        foreach ($result in $info) {
            if ($result.chapters.chapter.Count -ne $ChapterCount) {
                continue
            }

            $chapters = @()
            for ($index = 0; $index -lt $result.chapters.chapter.Count; $index++) {
                $chapters += New-Object PSObject -Property @{
                    Index = $index + 1
                    Time = [TimeSpan]$result.chapters.chapter[$index].time
                    Title = $result.chapters.chapter[$index].name
                }
            }

            $results += New-Object PSObject -Property @{
                Source = "ChaptersDb"
                ChaptersDbConfirmations = [int]$result.confirmations
                Title = $result.title
                Chapters = $chapters | Sort-Object Index
            }
        }

        Write-Verbose "Found $($chapters.Count) matches from ChaptersDb."
    }

    if ($UseBothServices -or 
            ($results.Count -eq 0 -and
             [String]::IsNullOrEmpty($TagChimpApiKey) -eq $false)) {
        $tagChimpApi = "https://www.tagchimp.com/ape/search.php"
        if ($ChapterCount -eq $null) {
            $count = "X"
        } else {
            $count = $ChapterCount
        }
        $tagChimpRequest = "{0}?token={1}&type=search&title={2}&totalChapters={3}" -f $tagChimpApi, $TagChimpApiKey, $escapedTitle, $count

        $response = Invoke-WebRequest -Uri $tagChimpRequest -Method GET -Headers @{ ApiKey = $TagChimpApiKey }
        if ($response.StatusCode -ne 200) {
            Write-Error "Request unsuccessful. ($($response.StatusCode), $($response.StatusDescription))"
            return
        }

        $info = ([xml]($response.Content)).items.movie
        Write-Verbose "Found $($info.Count) matches from tagChimp."
        $results += foreach ($result in $info) {
            $chapters = @()
            foreach ($chapter in $result.movieChapters.chapter) {
                $chapters += New-Object PSObject -Property @{
                    Index = [int]$chapter.chapterNumber
                    Time = [TimeSpan]$chapter.chapterTime
                    Title = $chapter.chapterTitle
                }
            }

            New-Object PSObject -Property @{
                Source = "tagChimp"
                ChaptersDbConfirmations = $null
                Title = $result.movieTags.Info.movieTitle
                Chapters = $chapters | Sort-Object Index
            }
        }
    }


    if ($BestResult) {
        return @($results | sort Confirmations -Descending)[0]
    } else {
        return $results
    }
}
