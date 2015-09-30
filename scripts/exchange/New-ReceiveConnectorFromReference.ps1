[CmdletBinding(
        SupportsShouldProcess=$true,
        ConfirmImpact="High")]
param (
        [Parameter(Mandatory=$true,
            ParameterSetName="SourceConnector")]
        [ValidateNotNullOrEmpty()]
        [object]
        $SourceReceiveConnector,

        [Parameter(Mandatory=$true,
            ParameterSetName="SourceServer")]
        [ValidateNotNullOrEmpty()]
        [string]
        $SourceServer,

        [Parameter(Mandatory=$true,
            ParameterSetName="SourceServer")]
        [ValidateNotNullOrEmpty()]
        [string]
        $ConnectorName,

        [Parameter(Mandatory=$true)]
        [ValidateNotNullOrEmpty()]
        [string]
        $Server,

        [Parameter(Mandatory=$false)]
        [ValidateNotNullOrEmpty()]
        [string]
        $NewConnectorName,

        [switch]
        $ShowCommand = $false
        )

BEGIN {
    $getCommand = "Get-ReceiveConnector"
    $setCommand = "Set-ReceiveConnector"
    $newCommand = "New-ReceiveConnector"
}

PROCESS {
    if ($SourceReceiveConnector -ne $null) {
        $source = $SourceReceiveConnector
        $ConnectorName = $SourceReceiveConnector.Name
        $SourceServer = $SourceReceiveConnector.Server
    } else {
        $source = & $getCommand -Identity $SourceServer\$ConnectorName -ErrorAction SilentlyContinue
    }

    if ([String]::IsNullOrEmpty($NewConnectorName)) {
        $NewConnectorName = $ConnectorName
        $dest = & $getCommand $Server\$ConnectorName -ErrorAction SilentlyContinue
    }

    if ($source -eq $null) {
        throw "Source connector must exist"
    }

    $sourceProps = Get-Member -InputObject $source -MemberType Property | Select -Expand Name
    if ($dest -eq $null) {
        $cmdletParams = (Get-Command -Name $newCommand -CommandType Function | Select -Expand Parameters).Keys
    } else {
        $cmdletParams = (Get-Command -Name $setCommand -CommandType Function | Select -Expand Parameters).Keys
    }

    $params = @()
    foreach ($prop in $sourceProps) {
        if ( @("Identity", "Server", "Name") -contains $prop) { continue }
        if ($cmdletParams -contains $prop) {
            if (($source.$prop -replace $SourceServer, $Server) -eq $dest.$prop) { continue }

            $equal = $true
            if (($source.$prop).Count) {
                foreach ($m in $source.$prop) {
                    if ($dest.$prop -notcontains $m) {
                        $equal = $false
                        break
                    }
                }
            } else {
                if ($source.$prop -ne $dest.$prop) {
                    $equal = $false
                }
            }
            if (!$equal) {
                Write-Verbose "$prop on $SourceServer and $Server are different"
                $params += $prop
            }
        }
    }

    if ($params.Count -eq 0) {
        Write-Verbose "All properties match between source and destination"
        return
    }

    if ($dest -eq $null) {
        $command = $newCommand
        $cmd = "$command -Server $Server -Name '$NewConnectorName'"
    } else {
        $command = $setCommand
        $cmd = "$command -Identity '$Server\$NewConnectorName'"
    }

    foreach ($param in $params) {
        if ([String]::IsNullOrEmpty($source.$param)) { continue }

        if ($source.$param -is [Microsoft.Exchange.Data.ByteQuantifiedSize]) {
            $val = ($source.$param).ToBytes()
        } elseif ($source.$param -is [Microsoft.Exchange.Data.Unlimited[Int]]) {
            $val = $source.$param
        } elseif ($source.$param -is [Microsoft.Exchange.Data.EnhancedTimeSpan]) {
            $val = ($source.$param).ToString()
        } elseif ($source.$param -is [Int]) {
            $val = $source.$param
        } elseif ($source.$param -is [Boolean]) {
            $val = '${0}' -f $source.$param
        } elseif ($param -eq "AuthMechanism") {
            $val = ($source.$param.ToString().Split(",") | % { $_.Trim() }) -join ", "
        } elseif ($param -eq "Bindings" -or $param -eq "RemoteIPRanges") {
            $val = ($source.$param | % { '"{0}"' -f $_ }) -join ", "
        } else {
            if ($source.$param -match $SourceServer) {
                $val = '"{0}"' -f ($source.$param -replace $SourceServer, $Server)
            } else {
                $val = '"{0}"' -f $source.$param
            }
        }

        $cmd += " ```n    -{0} {1}" -f $param, $val
    }

    if ($ShowCommand) {
        Write-Host $cmd
    }

    if ($PSCmdlet.ShouldProcess("$Server\$ConnectorName", $command)) {
        Invoke-Expression $cmd
    }
}
