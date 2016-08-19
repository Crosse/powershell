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
    Gets the current geographic location based on the compters's IP address.

    .DESCRIPTION
    Gets the current geographic location of the computer based on the
    computer's IP address using freegeoip.net. (This is only as good as the
    geolocation database.)

    .INPUTS
    None. You cannot pipe data into this cmdlet.

    .OUTPUTS
    System.Management.Automation.PSCustomObject.  Get-GeoLocation returns
    various details about the current location as based on the computer's IP
    address.

    .EXAMPLE
    PS C:\> Get-GeoLocation

    Ip          : 203.0.113.81
    CountryCode : US
    CountryName : United States
    RegionCode  : VA
    RegionName  : Virginia
    City        : Harrisonburg
    ZipCode     : 22807
    Latitude    : 38.4409
    Longitude   : -78.8742
    MetroCode   : 569

#>
function Get-GeoLocation {
    [CmdletBinding()]
    param (
            [Parameter(Mandatory=$false,
                ValueFromPipeline=$true)]
            [ValidateNotNullOrEmpty()]
            [string]
            # The IP address to geolocate.
            $IPAddress
          )

    if ($PSBoundParameters.ContainsKey("IPAddress") -eq $false) {
        Write-Verbose "Getting public IP address"
        $IPAddress = Get-PublicIPAddress
        if ([String]::IsNullOrEmpty($ip)) {
            throw "Unable to determine public IP address"
        }
    }

    Write-Verbose "Querying location data for IP address $IPAddress"
    $VerbosePreference = "SilentlyContinue"
    $results = Invoke-RestMethod -UseBasicParsing "https://freegeoip.net/json/$IPAddress"

    # This is only to make it look prettier.
    return New-Object -TypeName PSObject -Property @{
        IPAddress   = $ip
        CountryCode = $results.country_code
        CountryName = $results.country_name
        RegionCode  = $results.region_code
        RegionName  = $results.region_name
        City        = $results.city
        ZipCode     = $results.zip_code
        TimeZone    = $results.time_zone
        Latitude    = $results.latitude
        Longitude   = $results.longitude
        MetroCode   = $results.metro_code
    }
}

