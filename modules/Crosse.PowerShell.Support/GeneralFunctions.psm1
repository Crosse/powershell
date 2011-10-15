################################################################################
#
# Copyright (c) 2011 Seth Wright <seth@crosse.org>
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

# Functions going in this file should be readable in one page-length.  Otherwise,
# consider putting them in a separate module.

################################################################################
<#
    .SYNOPSIS
    Downloads a file from a remote URL.

    .DESCRIPTION
    Downloads a file from a remote URL.  The file can either be saved or
    converted to a System.String.

    .INPUTS
    System.String.  The URL to retrieve can be passed via the pipeline.

    .OUTPUTS
    If -AsString is True, then return the contents of the remote file as
    a System.String.  Otherwise, nothing is returned.

    .EXAMPLE
    C:\PS> Get-RemoteFile -Url http://checkip.dyndns.org -AsString
    <html><head><title>Current IP Check</title></head><body>Current IP Address: 203.0.113.81</body></html>
#>
################################################################################
function Get-RemoteFile {
    param (
            [Parameter(Mandatory=$true,
                ValueFromPipeline=$true)]
            [string]
            # The URL to retrieve.
            $Url,

            [switch]
            # Whether to return save the file to disk or return
            # the data as a System.String. The default is False.
            $AsString
          )

    # Attempt to download the file if a URL was specified.
    $wc = New-Object Net.WebClient
        if ($AsString) {
            $file = $wc.DownloadString($Url)
                Write-Output $file
        } else {
            $FileName = [System.IO.Path]::Combine((Get-Location).Path, $Url.Substring($Url.LastIndexOf("/") + 1))
                Write-Host "Saving to file $FileName"
                $wc.DownloadFile($Url, $FileName)
        }
    $wc.Dispose()
}


################################################################################
<#
    .SYNOPSIS
    Generates a new Globally-Unique Identifier (GUID).

    .DESCRIPTION
    Generates a new Globally-Unique Identifier (GUID).

    .INPUTS
    None.  You cannot pipe objects to New-Guid.

    .OUTPUTS
    System.Guid.

    .EXAMPLE
    C:\PS> New-Guid

    Guid
    ----
    f6bf4469-7419-4356-aac9-074de4e00a17
#>
################################################################################
function New-Guid {
    return [System.Guid]::NewGuid()
}

function Get-Uptime {
    param (
            [switch]
            $AsObject
          )

    # 2:31PM  up 5 days, 15:59, 2 users, load averages: 6.46, 6.65, 6.98
    $prop = Get-WmiObject Win32_OperatingSystem -Property LastBootUpTime
    $lastBootUpTime = $prop.ConvertToDateTime($prop.LastBootUpTime)
    $uptime = (Get-Date) - $lastBootUpTime
    $now = Get-Date
    $unixString = "{0,7} up {1} days, {2}:{3:00}" -f
                    $now.ToString("h:mmtt"),
                    $uptime.Days,
                    $uptime.Hours,
                    $uptime.Minutes

    if (!$AsObject) {
        $unixString
    } else {
        New-Object PSObject -Property @{
            LastBootUpTime     = $lastBootUpTime
            CurrentTime        = $now
            UnixStyleString    = $unixString
            UptimeTimeSpan     = $uptime
        }
    }
}
