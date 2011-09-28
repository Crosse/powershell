################################################################################
#
# $URL$
# $Author$
# $Date$
# $Rev$
#
# Copyright (c) 2009,2010 Seth Wright <wrightst@jmu.edu>
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

function Get-HexDump {
    param (
            [Parameter(Mandatory=$true, ValueFromPipeline=$true)]
            [string]
            # The path to the file to view in hex.
            $Path,

            [int]
            # The width of the hex dump. The default is 10.
            $Width=10,

            [int]
            # The number of bytes to decode. The default is -1.
            $Bytes=-1
          );

    $OFS=""
    Get-Content -Encoding byte $Path -ReadCount $Width -totalcount $Bytes | % {
        $byte = $_
            if (($byte -eq 0).count -ne $Width)
            {
                $hex = $byte | % {
                    " " + ("{0:x}" -f $_).PadLeft(2,"0")}
                $char = $byte | % {
                    if ([char]::IsLetterOrDigit($_))
                    { [char] $_ } else { "." }}
                "$hex $char"
            }
    }

    <#
        .SYNOPSIS
        Gets the contents of a file and displays it as a hex dump.

        .DESCRIPTION
        Gets the contents of a file and displays it as a hex dump.
        The width and length of the dump can be configured.

        .INPUTS
        A System.String describing the path to a file can be passed
        via the pipeline.

        .OUTPUTS
        System.String.  Get-HexDump returns a formatted hex dump of
        the file specified.

        .EXAMPLE
        C:\PS> Get-HexDump $Env:WINDIR\explorer.exe -Width 15 -Bytes 150
        4d 5a 90 00 03 00 00 00 04 00 00 00 ff ff 00 MZ..........ÿÿ.
        00 b8 00 00 00 00 00 00 00 40 00 00 00 00 00 ...............
        e8 00 00 00 0e 1f ba 0e 00 b4 09 cd 21 b8 01 è.....º....Í...
        4c cd 21 54 68 69 73 20 70 72 6f 67 72 61 6d LÍ.This.program
        20 63 61 6e 6e 6f 74 20 62 65 20 72 75 6e 20 .cannot.be.run.
        69 6e 20 44 4f 53 20 6d 6f 64 65 2e 0d 0d 0a in.DOS.mode....
        24 00 00 00 00 00 00 00 e2 b1 38 08 a6 d0 56 ........â.8..ÐV
        5b a6 d0 56 5b a6 d0 56 5b af a8 d2 5b ef d0 ..ÐV..ÐV...Ò.ïÐ
    #>

}

Set-Alias hd Get-HexDump

function ConvertTo-Base64 {
    [CmdletBinding()]
    param
        (
         [Parameter(ValueFromPipeline=$true)]
         # The string to convert.
         $InputObject
        );

    BEGIN {
        if ($InputObject -ne $null -and $InputObject.GetType() -eq [Object[]]) {
            $strings = New-Object System.Collections.ArrayList
        }
    }
    PROCESS {
        Write-Verbose "InputObject is a $($InputObject.GetType())"

        if ($InputObject.GetType() -eq [Object[]]) {
            foreach ($string in $InputObject) {
                $strings.Add($string) | Out-Null
            }
        } elseif ($InputObject.GetType() -eq [System.IO.FileInfo]) {
            $bytes = [System.IO.File]::ReadAllBytes($InputObject)
            [System.Convert]::ToBase64String($bytes, "InsertLineBreaks");
        }
    }
    END {
        if ($InputObject.GetType() -eq [Object[]]) {
            $joined = [String]::Join("`n", $strings.ToArray());
            $bytes  = [System.Text.Encoding]::UTF8.GetBytes($joined);
            [System.Convert]::ToBase64String($bytes, "InsertLineBreaks");
        }
    }

    <#
        .SYNOPSIS
        Converts a string to its base-64-encoded value.

        .DESCRIPTION
        Converts a string to its base-64-encoded value.

        .INPUTS
        None.  You cannot pipe objects to ConvertTo-Base64.

        .OUTPUTS
        System.String.  ConvertTo-Base64 returns the base-64-encoded value
        of the input string.

        .EXAMPLE
        C:\PS> ConvertTo-Base64 "Test String"
        VGVzdCBTdHJpbmc=
    #>
}

function ConvertFrom-Base64 {
        param (
            [Parameter(ValueFromPipeline=$true)]
            [String[]]
            # The string to convert.
            $InputObject,

            [String]
            $OutputFile
          );

    BEGIN {
        if ($OutputFile -notlike "*\*") {
            $OutputFile = "$pwd" + "\" + "$OutputFile"
        } elseif ($OutputFile -like ".\*") {
            $OutputFile = $OutputFile -replace "^\.",$pwd.Path
        } elseif ($OutputFile -like "..\*") {
            $OutputFile = $OutputFile -replace "^\.\.",$(get-item $pwd).Parent.FullName
        } else {
            throw "Cannot resolve path!"
        }

        $strings = New-Object System.Collections.ArrayList
    }
    PROCESS {
        foreach ($string in $InputObject) {
            $strings.Add($string) | Out-Null
        }
    }
    END {
        $joined = [String]::Join("`n", $strings.ToArray())
        $bytes  = [System.Convert]::FromBase64String($joined)
        if ($OutputFile) {
            [System.IO.File]::WriteAllBytes($OutputFile, $bytes)
        } else {
            [System.Text.Encoding]::UTF8.GetString($bytes);
        }
    }

    <#
        .SYNOPSIS
        Converts a string from base-64 to its unencoded value.

        .DESCRIPTION
        Converts a string from base-64 to its unencoded value.

        .INPUTS
        None.  You cannot pipe objects to ConvertFrom-Base64.

        .OUTPUTS
        System.String.  ConvertFrom-Base64 returns the original, unencoded
        value of a base-64-encoded string.

        .EXAMPLE
        C:\PS> ConvertFrom-Base64 "VGVzdCBTdHJpbmc="
        Test String
    #>
};

function Get-EventLogSummary {
    param (
            [Parameter(Mandatory=$false,
                ValueFromPipeline=$true)]
            [string]
            # Gets events from the event logs on the specified computer.
            $ComputerName,

            [Parameter(Mandatory=$true)]
            [string[]]
            # Gets events from the specified event logs. Enter the event log
            # names in a comma-separated list. Wildcards are permitted
            $LogName,

            [Parameter(Mandatory=$false)]
            [DateTime]
            # Gets only the events that occur after the specified date and
            # time. Enter a DateTime object, such as the one returned by the
            # Get-Date cmdlet.
            $StartTime,

            [Parameter(Mandatory=$false)]
            [DateTime]
            # Gets only the events that occur before the specified date and
            # time. Enter a DateTime object, such as the one returned by the
            # Get-Date cmdlet.
            $EndTime,

            [Parameter(Mandatory=$false)]
            [Int32[]]
            # Gets only the events that correspond to the specified log
            # level(s).
            $Level
        )

    BEGIN {
        $eventParams = @{ LogName=$LogName }
        if ($StartTime -ne $null) {
            $eventParams.Add('StartTime', $StartTime)
        }
        if ($EndTime -ne $Null) {
            $eventParams.Add('EndTime', $EndTime)
        }

        if ($Level -ne $null) {
            $eventParams.Add('Level', $Level)
        }
    }
    PROCESS {
        if ([String]::IsNullOrEmpty($ComputerName) -eq $false) {
            $logs = Get-WinEvent -ComputerName $ComputerName -FilterHashTable $eventParams
        } else {
            $logs = Get-WinEvent -FilterHashTable $eventParams
        }

        if ($logs -eq $null) {
            return
        }

        $retval = New-Object System.Collections.HashTable

        foreach ($event in $logs) {
            $dedupid = [String]::Format("{0}_{1}", $event.ProviderName , $event.Id)

            if ($retval.Contains($dedupid)) {
                $retval[$dedupid].Count++
                if ($retval[$dedupid].FirstTime -gt $event.TimeCreated) {
                    $retval[$dedupid].FirstTime = $event.TimeCreated
                }
                if ($retval[$dedupid].LastTime -lt $event.TimeCreated) {
                    $retval[$dedupid].LastTime = $event.TimeCreated
                }
            } else {
                $defaultProperties = @('Count','ProviderName','Id','LevelDisplayName','SampleMessage')
                $defaultDisplayPropertySet =
                    New-Object System.Management.Automation.PSPropertySet(
                            'DefaultDisplayPropertySet',
                            [string[]]$defaultProperties)

                $PSStandardMembers =
                    [System.Management.Automation.PSMemberInfo[]]@($defaultDisplayPropertySet)

                $obj = New-Object PSObject -Property @{
                    Count               = 1
                    ProviderName        = $event.ProviderName
                    Id                  = $event.Id
                    Level               = $event.Level
                    LevelDisplayName    = $event.LevelDisplayName
                    LogName             = $event.LogName
                    MachineName         = $event.MachineName
                    TaskDisplayName     = $event.TaskDisplayName
                    FirstTime           = $event.TimeCreated
                    LastTime            = $event.TimeCreated
                    SampleMessage       = $event.Message
                }

                Add-Member -InputObject $obj -MemberType MemberSet `
                                             -Name PSStandardMembers `
                                             -Value $PSStandardMembers

            }
            $retval[$dedupid] = $obj
        }

        $retval.GetEnumerator() | % { $_.Value }
    }
}


function Get-RemoteFile {
    param (
            [Parameter(Mandatory=$true, ValueFromPipeline=$true)]
            [string]
            # The URL to retrieve.
            $Url,

            [switch]
            # Whether to return save the file to disk or return
            # the data as a System.String. The default is False.
            $AsString=$false
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

    <#
        .SYNOPSIS
        Downloads a file from a remote URL.

        .DESCRIPTION
        Downloads a file from a remote URL.  The file can either be saved or
        converted to a System.String.

        .INPUTS
        The URL to retrieve can be passed via the pipeline.

        .OUTPUTS
        If -AsString is True, then return the contents of the remote file as
        a System.String.  Otherwise, nothing is returned.

        .EXAMPLE
        C:\PS> Get-RemoteFile -Url http://checkip.dyndns.org -AsString
        <html><head><title>Current IP Check</title></head><body>Current IP Address: 134.126.39.81</body></html>
    #>
}


function New-Guid {
    return [System.Guid]::NewGuid()

    <#
        .SYNOPSIS
        Generates a new Globally-Unique Identifier (GUID).

        .DESCRIPTION
        Generates a new Globally-Unique Identifier (GUID).

        .INPUTS
        None.  You cannot pipe objects to New-Guid.

        .OUTPUTS
        Returns a System.Guid.

        .EXAMPLE
        C:\PS> New-Guid

        Guid
        ----
        f6bf4469-7419-4356-aac9-074de4e00a17
    #>

}


function Set-Encoding {
    [CmdletBinding()]
        param (
                [Parameter(Position=0,Mandatory=$true,ValueFromPipeline=$true)]
                [ValidateScript({
                    if((resolve-path $_).Provider.Name -ne "FileSystem") {
                        throw "Specified Path is not in the FileSystem: '$_'"
                    }
                    return $true
                    })]
                [Alias("Fullname","Path")]
                [string]
                # The path to the file to be re-encoded.
                $FilePath,

                [switch]
                # Outputs the file in Unicode format.
                $Unicode,

                [switch]
                # Outputs the file in UTF7 format.
                $UTF7,

                [switch]
                # Outputs the file in UTF8 format.
                $UTF8,

                [switch]
                # Outputs the file in UTF32 format.
                $UTF32,

                [switch]
                # Outputs the file in ASCII format.
                $ASCII,

                [switch]
                # Outputs the file in BigEndianUnicode format.
                $BigEndianUnicode,

                [switch]
                # Uses the encoding of the system's current ANSI code page.
                $Default,

                [switch]
                # Uses the current original equipment manufacturer code page identifier for the operating system.
                $OEM
             )

        BEGIN {
            $Encoding = ""
            switch(
                    $Unicode,
                    $UTF7,
                    $UTF8,
                    $UTF32,
                    $ASCII,
                    $BigEndianUnicode,
                    $Default,
                    $OEM
                ) {
                $Unicode { $Encoding = "Unicode" }
                $UTF7 { $Encoding = "UTF7" }
                $UTF8 { $Encoding = "UTF8" }
                $UTF32 { $Encoding = "UTF32" }
                $ASCII { $Encoding = "ASCII" }
                $BigEndianUnicode { $Encoding = "BigEndianUnicode" }
                $Default { $Encoding = "Default" }
                $OEM { $Encoding = "OEM" }
            }

        }

    PROCESS {
        (Get-Content $FilePath) | Out-File -FilePath $FilePath -Encoding $Encoding -Force
    }

    <#
        .SYNOPSIS
        Takes a Script file or any other text file into memory
        and Re-Encodes it in the format specified.

        .EXAMPLE
        ls *.ps1 | Set-Encoding -ASCII

        .DESCRIPTION
        Written to provide an easy method to perform easy batch
        encoding, calls on the command Out-File with the -Encoding
        parameter and the -Force switch. Primarily to fix UnknownError
        status received when trying to sign non-ascii format files with
        digital signatures. Don't use on your MP3's or other non-text
        files :)
#>

}

function Get-CurrentWeather {
    param (
            [Parameter(Mandatory=$true,
                ValueFromPipeline=$true)]
            [ValidateNotNullOrEmpty()]
            [string]
            # The ZIP Code, City, Personal Weather Station, or Airport Code to
            # look up.
            $Identity
          )

    $api = "http://api.wunderground.com/auto/wui/geo/WXCurrentObXML/index.xml?query="
    $wc = New-Object System.Net.WebClient

    $query = $api + $Identity
    [xml]$weather = $wc.DownloadString($query)

    $weather.current_observation

    <#
        .SYNOPSIS
        Requests the current weather observations from Weather Underground
        for a particular airport, ZIP code, or personal weather station.

        .DESCRIPTION
        Requests the current weather observations from Weather Underground
        for a particular airport, ZIP code, or personal weather station.
        Information on the API can be found at
        http://wiki.wunderground.com/index.php/API_-_XML

        .OUTPUTS
        Xml.XmlElement.  Get-CurrentWeather returns an XML fragment detailing
        the current weather conditions.

        .EXAMPLE
        PS C:\Users\wrightst\Desktop> Get-CurrentWeather '22801'


        credit                  : Weather Underground NOAA Weather Station
        credit_URL              : http://wunderground.com/
        termsofservice          : termsofservice
        image                   : image
        display_location        : display_location
        observation_location    : observation_location
        station_id              : KSHD
        observation_time        : Last Updated on February 1, 5:20 PM EST
        observation_time_rfc822 : Tue, 01 February 2011 22:20:00 GMT
        observation_epoch       : 1296598800
        local_time              : February 1, 5:27 PM EST
        local_time_rfc822       : Tue, 01 February 2011 22:27:26 GMT
        local_epoch             : 1296599246
        weather                 : Overcast
        temperature_string      : 46 F (8 C)
        temp_f                  : 46
        temp_c                  : 8
        relative_humidity       : 76%
        wind_string             : Calm
        wind_dir                : North
        wind_degrees            : 0
        wind_mph                : 0
        wind_gust_mph           :
        pressure_string         : 30.05 in (1018 mb)
        pressure_mb             : 1018
        pressure_in             : 30.05
        dewpoint_string         : 39 F (4 C)
        dewpoint_f              : 39
        dewpoint_c              : 4
        heat_index_string       : NA
        heat_index_f            : NA
        heat_index_c            : NA
        windchill_string        : NA
        windchill_f             : NA
        windchill_c             : NA
        visibility_mi           : 10.0
        visibility_km           : 16.1
        icons                   : icons
        icon_url_base           : http://icons-ecast.wxug.com/graphics/conds/
        icon_url_name           : .GIF
        icon                    : cloudy
        forecast_url            : http://www.wunderground.com/US/VA/Harrisonburg.html
        history_url             : http://www.wunderground.com/history/airport/KSHD/2011/2/1/DailyHistory.html
        ob_url                  : http://www.wunderground.com/cgi-bin/findweather/getForecast?query=38.25999832,-78.90000153

    #>
}
