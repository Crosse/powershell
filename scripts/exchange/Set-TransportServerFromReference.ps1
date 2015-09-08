[CmdletBinding(
        SupportsShouldProcess=$true,
        ConfirmImpact="High")]
param (
        [Parameter(Mandatory=$true)]
        [ValidateNotNullOrEmpty()]
        [string]
        $SourceServer,

        [Parameter(Mandatory=$true)]
        [ValidateNotNullOrEmpty()]
        [string]
        $Server,

        [switch]
        $ShowCommand = $false
        )

BEGIN {
    $getCommand = "Get-TransportServer"
    $setCommand = "Set-TransportServer"
}

PROCESS {
    $source = & $getCommand $SourceServer
    $new = & $getCommand $Server

    $sourceProps = Get-Member -InputObject $source -MemberType Property | Select -Expand Name
    $cmdletParams = (Get-Command -Name Set-TransportServer -CommandType Function | Select -Expand Parameters).Keys
    $params = @()
    foreach ($prop in $sourceProps) {
        if ($prop -eq "Identity") { continue }
        if ($cmdletParams -contains $prop) {
            if ($source.$prop -ne $new.$prop) {
                Write-Verbose "$prop on $SourceServer and $Server are different"
                $params += $prop
            }
        }
    }

    if ($params.Count -eq 0) {
        Write-Verbose "All properties match between source and destination"
        return
    }

    $cmd = "$setCommand -Identity $Server"
    foreach ($param in $params) {
        if ([String]::IsNullOrEmpty($source.$param)) { continue }

        if ($source.$param -is [Microsoft.Exchange.Data.ByteQuantifiedSize]) {
            $val = ($source.$param).ToBytes()
        } elseif ($source.$param -is [Microsoft.Exchange.Data.Unlimited[Microsoft.Exchange.Data.ByteQuantifiedSize]]) {
            if (!($source.$param).IsUnlimited) {
                $val = ($source.$param).Value.ToBytes()
            }
        } elseif ($source.$param -is [Microsoft.Exchange.Data.EnhancedTimeSpan]) {
            $val = ($source.$param).ToString()
        } elseif ($source.$param -is [Int]) {
            $val = $source.$param
        } elseif ($source.$param -is [Boolean]) {
            $val = '${0}' -f $source.$param
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

    if ($PSCmdlet.ShouldProcess($Server, "Set-TransportServer")) {
        Invoke-Expression $cmd
    }
}
