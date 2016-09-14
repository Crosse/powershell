function Get-GraphToken {
    [CmdletBinding()]
    param (
            [Parameter(Mandatory=$true)]
            [string]
            $ClientID,

            [Parameter(Mandatory=$true)]
            [string]
            $ClientSecret,

            [Parameter(Mandatory=$true)]
            [string]
            $TenantDomain,

            [Parameter(Mandatory=$false)]
            [string]
            $Resource = "https://graph.windows.net"
          )

          $loginURL = "https://login.microsoftonline.com"

          $url = "{0}/{1}/oauth2/token?api-version=1.0" -f $loginURL, $TenantDomain
          $body = @{
              grant_type      = "client_credentials"
              resource        = $Resource
              client_id       = $ClientID
              client_secret   = $ClientSecret
          }

    $OldVerbosePref = $VerbosePreference
    $VerbosePreference = "SilentlyContinue"
    $oauth = Invoke-RestMethod -Method Post -Uri $url -Body $body
    $VerbosePreference = $OldVerbosePref

    if ($oauth.access_token -ne $null) {
        return $oauth
    } else {
        Write-Error "Failed to retrieve OAuth2 token for $resource"
    }
}


function Save-GraphToken {
    [CmdletBinding()]
    param (
            [Parameter(Mandatory=$true,
                ValueFromPipeline=$true)]
            [object]
            $Token,

            [Parameter(Mandatory=$true)]
            [string]
            $Path
          )
    $Token | Export-Clixml -Depth 100 -Path $Path
}

function Restore-GraphToken {
    [CmdletBinding()]
    param (
            [Parameter(Mandatory=$true)]
            [string]
            $Path
          )
    $Token = Import-Clixml -Path $Path
    $expires_on = ([datetime]'1970-01-01 00:00:00').AddSeconds($Token.expires_on)
    if ($expires_on -lt (Get-Date)) {
        Write-Error "Token has expired. Use Get-GraphToken to retrieve a new token."
    } else {
        return $Token
    }
}

function Get-AzureReport {
    [CmdletBinding()]
    param (
            [Parameter(Mandatory=$true,
                ParameterSetName="Token")]
            [object]
            $Token,

            [Parameter(Mandatory=$true)]
            [string]
            $ReportName,

            [Parameter(Mandatory=$false)]
            [DateTime]
            $Start = (Get-Date).AddDays(-7),

            [Parameter(Mandatory=$false)]
            [DateTime]
            $End = (Get-Date),

            [Parameter(Mandatory=$false)]
            [switch]
            $AsJson= $false
          )

    # Constants
    $loginURL = "https://login.microsoftonline.com"
    $resource = "https://graph.windows.net"

    Write-Verbose "Searching for events between $Start and $End"

    $i = 1
    $headerParams = @{
        Authorization = "$($token.token_type) $($token.access_token)"
    }

    $eventTimeStart = "{0:s}Z" -f $Start.ToUniversalTime()
    $eventTimeEnd = "{0:s}Z" -f $End.ToUniversalTime()

    $url = "{0}/{1}/reports/{2}?api-version=beta" -f $resource, $TenantDomain, $ReportName
    $url = '{0}&$filter=eventTime ge {1} and eventTime lt {2}' -f $url, $eventTimeStart, $eventTimeEnd
    $url = [Uri]::EscapeUriString($url)

    # loop through each query page (1 through n)
    do {
        Write-Debug "Fetching page $i of data"

        $OldVerbosePref = $VerbosePreference
        $VerbosePreference = "SilentlyContinue"
        $myReport = (Invoke-WebRequest -UseBasicParsing -Headers $headerParams -Uri $url)
        $VerbosePreference = $OldVerbosePref

        $content = $myReport.Content
        if (!$AsJson) {
            $content = ($myReport.Content | ConvertFrom-Json).value
        }

        foreach ($event in $content) {
            Write-Output $event
        }

        $url = $content.'@odata.nextLink'
        $i++
    } while ($url -ne $null)
}
