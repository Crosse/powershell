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

function Get-GeoLocation {
    $wc = New-Object System.Net.WebClient
    [xml]$response = $wc.DownloadString("http://freegeoip.net/xml/")
    $wc.Dispose()

    return $response.Response
}

function Get-Weather {
    [CmdletBinding()]
    param (
            [Parameter(Mandatory=$false,
                ValueFromPipeline=$true)]
            [string]
            # The ZIP Code, City, Personal Weather Station, or Airport Code to
            # look up.
            $Location,

            [switch]
            $AsObject=$false
          )

    if ([String]::IsNulLOrEmpty($Location)) {
        $Location = (Get-GeoLocation).ZipCode
        if ($Location -eq $null) {
            Write-Error "Cannot dynamically determine Location."
        }
        Write-Verbose "Found location $Location"
    }

    $api = "http://api.wunderground.com/auto/wui/geo/WXCurrentObXML/index.xml?query="
    $wc = New-Object System.Net.WebClient
    $wc.Dispose()

    $query = $api + $Location
    [xml]$weather = $wc.DownloadString($query)

    if (!$AsObject) {
        $result = "For {0} it is currently {1} degrees, with {2} skies." -f
                        $weather.current_observation.observation_location.full,
                        $weather.current_observation.temp_f,
                        $weather.current_observation.weather.ToLower()
    } else {
        $result = $weather.current_observation
    }

    return $result

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
