function Get-ESVariableOrDefault {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory=$true)]
        [string]
        $Name,

        [Parameter(Mandatory=$false)]
        $DefaultValue
    )

    $val = Get-Variable -Name $Name -ValueOnly -ErrorAction SilentlyContinue
    return $(if ($val) { $val } else { $DefaultValue })
}

function ParseCatData {
    param (
        [Parameter(Mandatory=$true,
            ValueFromPipeline=$true)]
        [string]
        $Data
    )

    PROCESS {
        $trimmed = ($Data -split "`n") | % { ($_ -replace '\s+', ' ').Trim() }
        $headers = ($trimmed[0] -replace '\.', '_').Split()
        $headers = foreach ($hdr in $headers) {
            switch ($hdr) {
                "ip" { "IPAddress" }
                "prirep" { "PrimaryOrReplica" }
                "pri" { "PrimaryShards" }
                "rep" { "ReplicaShards" }
                default {
                    $chars = @('_') + $hdr.ToCharArray()
                    $newhdr = ""
                    for ($i = 0; $i -lt $chars.Count; $i++) {
                        if ($chars[$i] -eq '_') {
                            $i++
                            $newhdr += ([string]$chars[$i]).ToUpper()
                        } else {
                            $newhdr += [string]$chars[$i]
                        }
                    }
                    $newhdr
                }
            }
        }

        return $trimmed[1..($trimmed.Length-1)] | ConvertFrom-Csv -Delimiter ' ' -Header $headers
    }
}

function Invoke-ESCatEndpoint {
    [CmdletBinding()]
        param (
        [Parameter(Mandatory=$false)]
        [ValidateNotNullOrEmpty()]
        [string]
        $Hostname = (Get-ESVariableOrDefault -Name "ESHostname" -DefaultValue "localhost"),

        [Parameter(Mandatory=$false)]
        [ValidateRange(1,65535)]
        [Int32]
        $Port = 9200,

        [Parameter(Mandatory=$false)]
        [ValidateNotNullOrEmpty()]
        [ValidateScript({$_.StartsWith("/")})]
        [string]
        $UrlPrefix = "/",

        [switch]
        $UseSsl = (Get-ESVariableOrDefault -Name "ESUseSsl" -DefaultValue $false),

        [Parameter(Mandatory=$true)]
        [ValidateNotNullOrEmpty()]
        [string]
        $CatEndpoint,

        [switch]
        $SizesInBytes = $false,

        [Parameter(Mandatory=$false)]
        [string]
        $ObjectName
    )

    PROCESS {
        $uri = "http{0}://{1}:{2}{3}_cat/{4}/{5}?v" -f $(if ($UseSsl) { "s" }), $Hostname, $Port, $UrlPrefix, $CatEndpoint, $ObjectName
        if ($SizesInBytes) {
            $uri += "&bytes=b"
        }
        Write-Verbose "GET $uri"

        $result = $null
        try {
            # Invoke-WebRequest doesn't need to be verbose.
            $vp = $VerbosePreference
            $VerbosePreference = "SilentlyContinue"
            $resp = Invoke-WebRequest -Uri $uri
            $VerbosePreference = $vp

            $result = $resp.Content | ParseCatData
            
        } catch {
            if ($_ -isnot [string]) {
                Write-Error $_
            } else {
                if ($_.StartsWith('{')) {
                    $err = $_ | ConvertFrom-Json
                    Write-Error ("Unable to invoke _cat endpoint {0}: {1}" -f $CatEndpoint, $err.error.reason)
                    
                } else {
                    Write-Error ("Unable to invoke _cat endpoint {0}: {1}" -f $CatEndpoint, $_)
                }
            }
        }
        return $result
    }
}

function Get-ESAlias {
    [CmdletBinding()]
        param (
        [Parameter(Mandatory=$false)]
        [string]
        $Alias,

        [Parameter(Mandatory=$false)]
        [ValidateNotNullOrEmpty()]
        [string]
        $Hostname = (Get-ESVariableOrDefault -Name "ESHostname" -DefaultValue "localhost"),

        [Parameter(Mandatory=$false)]
        [ValidateRange(1,65535)]
        [Int32]
        $Port = (Get-ESVariableOrDefault -Name "ESPort" -DefaultValue 9200),

        [Parameter(Mandatory=$false)]
        [ValidateNotNullOrEmpty()]
        [ValidateScript({$_.StartsWith("/")})]
        [string]
        $UrlPrefix = (Get-ESVariableOrDefault -Name "ESUrlPrefix" -DefaultValue "/"),

        [switch]
        $UseSsl = (Get-ESVariableOrDefault -Name "ESUseSsl" -DefaultValue $false),

        [switch]
        $SizesInBytes = $true
    )

    PROCESS {
        $null = $PSBoundParameters.Remove("Alias")
        return Invoke-ESCatEndpoint -CatEndpoint "aliases" -ObjectName $Alias @PSBoundParameters
    }
}

function Get-ESSegment {
    [CmdletBinding()]
        param (
        [Parameter(Mandatory=$false)]
        [string]
        $Index,

        [Parameter(Mandatory=$false)]
        [ValidateNotNullOrEmpty()]
        [string]
        $Hostname = (Get-ESVariableOrDefault -Name "ESHostname" -DefaultValue "localhost"),

        [Parameter(Mandatory=$false)]
        [ValidateRange(1,65535)]
        [Int32]
        $Port = (Get-ESVariableOrDefault -Name "ESPort" -DefaultValue 9200),

        [Parameter(Mandatory=$false)]
        [ValidateNotNullOrEmpty()]
        [ValidateScript({$_.StartsWith("/")})]
        [string]
        $UrlPrefix = (Get-ESVariableOrDefault -Name "ESUrlPrefix" -DefaultValue "/"),

        [switch]
        $UseSsl = (Get-ESVariableOrDefault -Name "ESUseSsl" -DefaultValue $false),

        [switch]
        $SizesInBytes = $true
        )

    PROCESS {
        $null = $PSBoundParameters.Remove("Index")
        return Invoke-ESCatEndpoint -CatEndpoint "segments" -ObjectName $Index @PSBoundParameters
    }
}

function Get-ESIndex {
    [CmdletBinding()]
        param (
        [Parameter(Mandatory=$false)]
        [string]
        $Index,

        [Parameter(Mandatory=$false)]
        [ValidateNotNullOrEmpty()]
        [string]
        $Hostname = (Get-ESVariableOrDefault -Name "ESHostname" -DefaultValue "localhost"),

        [Parameter(Mandatory=$false)]
        [ValidateRange(1,65535)]
        [Int32]
        $Port = (Get-ESVariableOrDefault -Name "ESPort" -DefaultValue 9200),

        [Parameter(Mandatory=$false)]
        [ValidateNotNullOrEmpty()]
        [ValidateScript({$_.StartsWith("/")})]
        [string]
        $UrlPrefix = (Get-ESVariableOrDefault -Name "ESUrlPrefix" -DefaultValue "/"),

        [switch]
        $UseSsl = (Get-ESVariableOrDefault -Name "ESUseSsl" -DefaultValue $false),

        [switch]
        $SizesInBytes = $true
    )

    PROCESS {
        $null = $PSBoundParameters.Remove("Index")

        $segments = Get-ESSegment -Index $Index @PSBoundParameters | Group-Object -NoElement Index
        $indices = Invoke-ESCatEndpoint -CatEndpoint "indices" -ObjectName $Index @PSBoundParameters

        foreach ($idx in $indices) {
            $s = $segments | ? { $_.Name -eq $idx.Index }
            $totalShards = ([Int]$idx.PrimaryShards + ([Int]$idx.PrimaryShards * [Int]$idx.ReplicaShards))
            $idx | Add-Member -PassThru -MemberType NoteProperty -Name "Segments" -Value $s.Count
        }
    }
}


function Find-ESUnOptimizedIndex {
    [CmdletBinding()]
        param (
        [Parameter(Mandatory=$false)]
        [string]
        $Index,

        [Parameter(Mandatory=$false)]
        [ValidateNotNullOrEmpty()]
        [string]
        $Hostname = (Get-ESVariableOrDefault -Name "ESHostname" -DefaultValue "localhost"),

        [Parameter(Mandatory=$false)]
        [ValidateRange(1,65535)]
        [Int32]
        $Port = (Get-ESVariableOrDefault -Name "ESPort" -DefaultValue 9200),

        [Parameter(Mandatory=$false)]
        [ValidateNotNullOrEmpty()]
        [ValidateScript({$_.StartsWith("/")})]
        [string]
        $UrlPrefix = (Get-ESVariableOrDefault -Name "ESUrlPrefix" -DefaultValue "/"),

        [Parameter(Mandatory=$false)]
        [UInt32]
        $MaxSegmentsPerShard = 1,

        [switch]
        $UseSsl = (Get-ESVariableOrDefault -Name "ESUseSsl" -DefaultValue $false),

        [switch]
        $SizesInBytes = $true
        )

    PROCESS {
        $null = $PSBoundParameters.Remove("MaxSegmentsPerShard")
        if (!$PSBoundParameters.ContainsKey("Index") -and ![String]::IsNullOrEmpty($Index)) {
            $PSBoundParameters.Add("Index", $Index)
        }
        $indices = Get-ESIndex @PSBoundParameters

        foreach ($idx in $indices) {
            $totalShards = ([Int]$idx.PrimaryShards + ([Int]$idx.PrimaryShards * [Int]$idx.ReplicaShards))
            $desired = $MaxSegmentsPerShard * $totalShards

            if ($idx.Segments -gt $desired) {
                $idx | Add-Member -PassThru -NotePropertyMembers @{
                    TotalShards = $totalShards
                    DesiredSegments = $desired
                }
            }
        }
    }
}

function Remove-ESIndex {
    [CmdletBinding(
        SupportsShouldProcess=$true,
        ConfirmImpact="High")]
    param (
        [Parameter(Mandatory=$true,
            ValueFromPipeline=$true)]
        [ValidateNotNullOrEmpty()]
        [string]
        $Index,

        [Parameter(Mandatory=$false)]
        [ValidateNotNullOrEmpty()]
        [string]
        $Hostname = (Get-ESVariableOrDefault -Name "ESHostname" -DefaultValue "localhost"),

        [Parameter(Mandatory=$false)]
        [ValidateRange(1,65535)]
        [Int32]
        $Port = (Get-ESVariableOrDefault -Name "ESPort" -DefaultValue 9200),

        [Parameter(Mandatory=$false)]
        [ValidateNotNullOrEmpty()]
        [ValidateScript({$_.StartsWith("/")})]
        [string]
        $UrlPrefix = (Get-ESVariableOrDefault -Name "ESUrlPrefix" -DefaultValue "/"),

        [switch]
        $UseSsl = (Get-ESVariableOrDefault -Name "ESUseSsl" -DefaultValue $false)        
    )

    PROCESS {
        $uri = "http{0}://{1}:{2}{3}{4}" -f $(if ($UseSsl) { "s" }), $Hostname, $Port, $UrlPrefix, $Index
        Write-Verbose "Deleting $Index"

        if ($PSCmdlet.ShouldProcess($Index)) {
            try {
                $vp = $VerbosePreference
                $VerbosePreference = "SilentlyContinue"
                $resp = Invoke-WebRequest -Uri $uri -Method Delete
                $VerbosePreference = $vp

                $ack = $resp.Content | ConvertFrom-Json
                if (!$ack.acknowledged) {
                    throw "Delete request was not acknowledged"
                }
            } catch {
                $err = $_ | ConvertFrom-Json
                Write-Error ("Unable to delete index {0}: {1}" -f $Index, $err.error.reason)
            }
        }
    }
}

