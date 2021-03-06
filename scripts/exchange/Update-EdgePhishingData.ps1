################################################################################
#
# DESCRIPTION:  Updates the AntiPhishing data files on the edge transports
#
# Copyright (c) 2009-2015 Seth Wright <wrightst@jmu.edu>
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
[CmdletBinding()]param ()

$now = Get-Date
Start-Transcript "update_edgephishingdata_$($now.Year)-$($now.Month)-$($now.Day)-$($now.Hour).log"

$baseUrl = "http://svn.code.sf.net/p/aper/code/"
$dataPath = "\e$\Program Files\Microsoft\Exchange Server\V14\TransportRoles\Agents\AntiPhishing\bin\Data"
$files = @( "phishing_reply_addresses" )

$localReplyAddresses = "E:\Scripts\local_phishing_reply_addresses.txt"

Write-Output "baseUrl = `"$baseUrl`""
Write-Output "dataPath = `"$dataPath`""
Write-Output "files = `"$files`""
Write-Output "$pwd"
Write-Output ""

$wc = New-Object System.Net.WebClient

foreach ($file in $files) {
    $error.Clear()
    Write-Verbose "Downloading $file"
    $fileData = $wc.DownloadString("$($baseUrl)/$($file)") -split "`n"

    # Append local data
    if (Test-Path $localReplyAddresses) {
        Write-Verbose "Local reply addresses file found.  Appending to $file"
        $fileData += Get-Content $localReplyAddresses
    }

    # Write the file out locally first.
    Write-Verbose "Saving $file to disk"
    Out-File -FilePath $file -InputObject $fileData -Encoding ASCII


    if ($fileData -and [String]::IsNullOrEmpty($error[0])) {
        for ($i = 1; $i -le 4; $i++) {
            $server = "it-exet$($i).jmu.edu"
            $remotePath = "\\$($server)$($dataPath)\$($file)"

            Write-Verbose "Pushing $file to $server"
            $error.Clear()
            Copy-Item $file $remotePath

            if (![String]::IsNullOrEmpty($error[0])) {
                Write-Output "Error copying $file to $server"
            }
        }
    } else {
        Write-Output "Error downloading file: $($error[0].ToString())"
    }
}
Stop-Transcript
