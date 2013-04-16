#Requires -Version 2.0

function Invoke-ServerAudit {
    [CmdletBinding()]
    param(
            [Parameter(Mandatory=$true,
                ValueFromPipeline=$true)]
            [string]
            # The name of the computer to audit.
            $ComputerName,

            [switch]
            # Determines whether to run connectivity tests, such as
            # verifying ICMP ping works, as well as remote WMI and remote
            # PowerShell.
            $CheckConnectivity = $true,

            [switch]
            $CheckFirewall = $true,

            [switch]
            # Determines whether or not to run all audit tests, regardless
            # of the outcome of previous tests.
            $Force = $false
        )

    $result = New-Object PSObject

    if ($CheckConnectivity) {
        $status = Get-ComputerStatus -ComputerName $ComputerName -Force:$Force -Verbose:$VerbosePreference
        $result = Add-PropertiesToObject -Source $status -Destination $result
    }

    if ($Force -eq $false -and
            ($status.PingStatus -ne 'Success' -or
             $status.WmiStatus -ne 'Success' -or
             $status.PSRemotingStatus -ne 'Success')) {
        return $result
    }

    if ($CheckFirewall) {
        $fw = Get-FirewallStatus -ComputerName $ComputerName -Verbose:$VerbosePreference
        $result = Add-PropertiesToObject -Source $fw -Destination $result
    }


    return $result
}

function Get-FirewallStatus {
    [CmdletBinding()]
    param (
            [Parameter(Mandatory=$true)]
            [string]
            $ComputerName
          )

    $enabled = Invoke-Command -ComputerName $ComputerName {
        $fwMgr = New-Object -ComObject HNetCfg.FwMgr
        $enabled = $fwMgr.LocalPolicy.CurrentProfile.FirewallEnabled
        Remove-Variable fwMgr
        $enabled
    }

    if ($enabled) {
        Write-Verbose "$ComputerName : Firewall is enabled"
    } else {
        Write-Warning "Firewall is DISABLED on $ComputerName"
    }

    $result = New-Object PSObject -Property @{
        FirewallEnabled = $enabled
    }
    return $result
}

function Add-PropertiesToObject {
    param (
            [Parameter(Mandatory=$true)]
            [object]
            $Source,

            [Parameter(Mandatory=$true)]
            [object]
            $Destination
          )

    foreach ($prop in (Get-Member -InputObject $Source -MemberType NoteProperty)) {
        if ($Destination.$prop -ne $null) {
            Write-Warning "Encountered two properties with the same name: $prop"
        }
        $Destination = Add-Member -InputObject $Destination -PassThru `
                        -MemberType NoteProperty -Name $prop.Name `
                        -Value $Source.$($prop.Name)
    }

    return $Destination
}
