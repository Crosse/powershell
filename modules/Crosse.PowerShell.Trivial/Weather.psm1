################################################################################
#
# Copyright (c) 2016 Seth Wright <seth@crosse.org>
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

<#
    .SYNOPSIS
    Requests the current weather observations from Weather Underground for a
    particular airport, ZIP code, or personal weather station.

    .DESCRIPTION
    Requests the current weather observations from Weather Underground for a
    particular airport, ZIP code, or personal weather station.  Information on
    the API can be found at http://wiki.wunderground.com/index.php/API_-_XML.
    If no Location is given, Get-Weather will use the Get-GeoLocation cmdlet
    to get the current location and use that.

    .INPUTS
    None.  You cannot pipe data into this cmdlet.

    .OUTPUTS
    Xml.XmlElement.  Get-CurrentWeather returns an XML fragment detailing the
    current weather conditions.

    .EXAMPLE
    PS C:\> Get-Weather
    For Staunton, Virginia it is currently 66 degrees, with clear skies.

    The above example tells Get-Weather to use the current location as returned
    by the Get-GeoLocation cmdlet and displays the weather in an easy-to-read
    format.

    .EXAMPLE
    PS C:\> Get-Weather -AsObject -Location 77536

    credit                  : Weather Underground NOAA Weather Station
    credit_URL              : http://wunderground.com/
    termsofservice          : termsofservice
    image                   : image
    display_location        : display_location
    observation_location    : observation_location
    station_id              : KEFD
    observation_time        : Last Updated on October 15, 3:50 PM CDT
    observation_time_rfc822 : Sat, 15 Oct 2011 20:50:00 GMT
    observation_epoch       : 1318711800
    local_time              : October 15, 4:00 PM CDT
    local_time_rfc822       : Sat, 15 Oct 2011 21:00:00 GMT
    local_epoch             : 1318712400
    weather                 : Partly Cloudy
    temperature_string      : 82 F (28 C)
    temp_f                  : 82
    temp_c                  : 28
    relative_humidity       : 48%
    wind_string             : From the ESE at 4 MPH
    wind_dir                : ESE
    wind_degrees            : 120
    wind_mph                : 4
    wind_gust_mph           :
    pressure_string         : 30.05 in (1018 mb)
    pressure_mb             : 1018
    pressure_in             : 30.05
    dewpoint_string         : 61 F (16 C)
    dewpoint_f              : 61
    dewpoint_c              : 16
    heat_index_string       : 83 F (28 C)
    heat_index_f            : 83
    heat_index_c            : 28
    windchill_string        : NA
    windchill_f             : NA
    windchill_c             : NA
    visibility_mi           : 10.0
    visibility_km           : 16.1
    icons                   : icons
    icon_url_base           : http://icons-ak.wxug.com/graphics/conds/
    icon_url_name           : .GIF
    icon                    : partlycloudy
    forecast_url            : http://www.wunderground.com/US/TX/Deer_Park.html
    history_url             : http://www.wunderground.com/history/airport/KEFD/2011/10/15/DailyHistory.html
    ob_url                  : http://www.wunderground.com/cgi-bin/findweather/getForecast?query=29.61000061,-95.16000366

    The above example tells Get-Weather to display the weather for ZIP code
    77356 as an object, which returns all the data that Weather Underground
    provides for the location.
#>
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
            # Whether to return the data in an easy-to-read format ($false) or
            # to return all the data received from Weather Underground.
            $AsObject = $false

          )

    if ([String]::IsNulLOrEmpty($Location)) {
        $Location = (Get-GeoLocation).ZipCode
        if ($Location -eq $null) {
            throw "Cannot dynamically determine Location."
        }
        Write-Verbose "Found location $Location"
    }

    $query = "http://api.wunderground.com/auto/wui/geo/WXCurrentObXML/index.xml?query={0}" -f $Location

    [xml]$weather = Invoke-RestMethod -UseBasicParsing -Uri $query

    if (!$AsObject) {
        $result = "For {0} it is currently {1} degrees, with {2} skies." -f
                        $weather.current_observation.observation_location.full,
                        $weather.current_observation.temp_f,
                        $weather.current_observation.weather.ToLower()
    } else {
        $result = $weather.current_observation
    }

    return $result
}
