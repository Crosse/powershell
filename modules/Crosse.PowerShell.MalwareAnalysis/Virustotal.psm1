function Send-ToVirusTotal {
    [CmdletBinding()]
    param (
            [Parameter(Mandatory=$true)]
            [ValidateNotNullOrEmpty()]
            [string]
            # An API Key from Virus Total.
            $ApiKey,

            [Parameter(Mandatory=$true)]
            [ValidateNotNullOrEmpty()]
            [string]
            # The full path to the file to submit.
            $FilePath,

            [Parameter(Mandatory=$false)]
            [ValidateNotNullOrEmpty()]
            [string]
            # The original name of the file, if it has been renamed.
            $OriginalFileName
          )
    BEGIN {
        $submitUri = "https://www.virustotal.com/vtapi/v2/file/scan"
        $ser = New-Object System.Web.Script.Serialization.JavaScriptSerializer
    }
    PROCESS {
        $contentType, $body = New-MultipartFormData -Fields @{"apikey" = $ApiKey} -Files @{$OriginalFileName = $FilePath}
        $body = [System.Text.Encoding]::UTF8.GetBytes($body)

        $wc = New-Object System.Net.WebClient
        $wc.Headers.Add("Content-Type", $contentType)
        Write-Verbose "Uploading data to Virus Total..."
        $response = $wc.UploadData($submitUri, "POST", $body)
        $json = [System.Text.Encoding]::ASCII.GetString($response)
        $json
        $obj = $ser.DeserializeObject($json)
        $obj
    }
}

function Get-VirusTotalReport {
    [CmdletBinding()]
    param (
            [Parameter(Mandatory=$true)]
            [ValidateNotNullOrEmpty()]
            [string]
            # An API Key from Virus Total.
            $ApiKey,

            [Parameter(Mandatory=$true)]
            [ValidateNotNullOrEmpty()]
            [string]
            # The full path to the file to submit.
            $FilePath,

            [Parameter(Mandatory=$false)]
            [ValidateNotNullOrEmpty()]
            [string]
            # The original name of the file, if it has been renamed.
            $OriginalFileName
          )

    BEGIN {
        $ser = New-Object System.Web.Script.Serialization.JavaScriptSerializer
        $md5 = New-Object -TypeName System.Security.Cryptography.MD5CryptoServiceProvider
        $reportUri = "https://www.virustotal.com/vtapi/v2/file/report"
        $rescanUri = "https://www.virustotal.com/vtapi/v2/file/rescan"
    }

    PROCESS {
        try {
            # Get the MD5 hash and see if it already exists in Virus Total's database.
            $stream = [System.IO.File]::Open("$FilePath", "Open", "Read")
            $hash = [System.BitConverter]::ToString($md5.ComputeHash($stream))
            $hash = $hash.Replace('-', '').ToLower()
            Write-Verbose "File hash:  $hash"
        } catch {
            throw $_
        } finally {
            $stream.Close()
            $stream.Dispose()
        }

        $req = "apikey={0}&resource={1}" -f $apikey, $hash
        $req = [System.Text.Encoding]::UTF8.GetBytes($req)

        $wc = New-Object System.Net.WebClient
        $wc.Headers.Add("Content-Type", "application/x-www-form-urlencoded")
        $response = $wc.UploadData($reportUri, "POST", $req)
        $json = [System.Text.Encoding]::ASCII.GetString($response)
        $obj = $ser.DeserializeObject($json)

        switch ($obj.response_code) {
            0 { $results = Send-ToVirusTotal -ApiKey $ApiKey -FilePath $FilePath -FileName $FileName }
            1 {
                #$scans = New-Object PSObject
                #foreach ($key in $obj.scans.Keys) {
                    #$s = New-Object PSObject -Property @{
                                    #Detected    = $obj.scans[$key].detected
                                    #Version     = $obj.scans[$key].version
                                    #Result      = $obj.scans[$key].result
                                    #Update      = $obj.scans[$key].update
                                #}
                    #$scans = $scans | Add-Member -InputObject $scans -MemberType NoteProperty -Name $key -Value $s
                #}
                $results = New-Object PSObject -Property @{
                            ResponseCode    = "Found"
                            VerboseMessage  = $obj.verbose_msg
                            Resource        = $obj.resource
                            ScanId          = $obj.scan_id
                            MD5             = $obj.md5
                            SHA1            = $obj.sha1
                            SHA256          = $obj.sha256
                            ScanDate        = [DateTime]$obj.scan_date
                            Positives       = $obj.positives
                            Total           = $obj.total
                            Scans           = $obj.scans
                            Permalink       = $obj.permalink
                        }
            }
            -2 {
                $results = New-Object PSObject -Property @{
                            ResponseCode    = "QueuedForAnalysis"
                            VerboseMessage  = $obj.verbose_msg
                }
            }
        }
        $results
    }
}

function New-MultipartFormData {
    [CmdletBinding()]
    param (
            [Parameter(Mandatory=$false)]
            [System.Collections.HashTable]
            $Fields,

            [Parameter(Mandatory=$false)]
            [System.Collections.HashTable]
            $Files
          )

    $boundary = "--------" + [Guid]::NewGuid().ToString("n")
    $lines = @()
    foreach ($key in $Fields.Keys) {
        $lines += "--" + $boundary
        $lines += "Content-Disposition: form-data; name=`"$key`""
        $lines += ""
        $lines += $Fields[$key]
    }
    foreach ($key in $Files.Keys) {
        Write-Verbose "Finding file $($Files[$key])"
        if ((Test-Path $Files[$key]) -eq $false) {
            Write-Error "File not found:  $($Files[$key])"
            return
        }
        $f = Get-Item $Files[$key]
        $lines += "--" + $boundary
        $lines += 'Content-Disposition: form-data; name="file"; filename="{0}"' -f $key
        # TODO:  Actually *guess* the MIME type?
        $lines += 'Content-Type: application/octet-stream'
        $lines += ""
        $lines += Get-Content -ReadCount 0 -LiteralPath $f
    }
    $lines += "--" + $boundary + "--"
    $lines += ""
    $body = $lines -join "`r`n"
    $contentType = "multipart/form-data; boundary=$boundary"
    return @($contentType, $body)
}
