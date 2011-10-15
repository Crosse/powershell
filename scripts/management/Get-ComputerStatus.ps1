[CmdletBinding()]
param (
        [Parameter(Mandatory=$true,
            ValueFromPipeline=$true)]
        [string]
        $ComputerName
      )

BEGIN {
    $ping = New-Object System.Net.NetworkInformation.Ping
}

PROCESS {
    $result = New-Object PSObject
    Add-Member -InputObject $result -MemberType NoteProperty -Name "DnsStatus" -Value "NotFound"
    Add-Member -InputObject $result -MemberType NoteProperty -Name "HostName" -Value $ComputerName
    Add-Member -InputObject $result -MemberType NoteProperty -Name "PingStatus" -Value $null
    Add-Member -InputObject $result -MemberType NoteProperty -Name "PingRoundTripTime" -Value $null
    Add-Member -InputObject $result -MemberType NoteProperty -Name "WmiStatus" -Value $null
    Add-Member -InputObject $result -MemberType NoteProperty -Name "Manufacturer" -Value $null
    Add-Member -InputObject $result -MemberType NoteProperty -Name "Model" -Value $null
    Add-Member -InputObject $result -MemberType NoteProperty -Name "OSVersion" -Value $null
    Add-Member -InputObject $result -MemberType NoteProperty -Name "OSName" -Value $null
    Add-Member -InputObject $result -MemberType NoteProperty -Name "PSRemotingStatus" -Value $null
    Add-Member -InputObject $result -MemberType NoteProperty -Name "PSRemotingError" -Value $null
    
    try {
        $dns = [System.Net.Dns]::GetHostEntry($ComputerName)
        $result.DnsStatus = "Found"
        $result.HostName = $dns.HostName
    } catch {
        Write-Verbose "$ComputerName : Name not found in DNS."
        return $result
    }

    $pingResult = $ping.Send($ComputerName, 1000)
    if ($pingResult.Status -ne 'Success') {
        Write-Verbose "$ComputerName : Did not respond to ping."
        return $result
    } else {
        $result.PingStatus = $pingResult.Status
        $result.PingRoundTripTime = $pingResult.RoundTripTime
    }

    $error.Clear()
    $computerSystem = Get-WmiObject -ComputerName $ComputerName -Class Win32_ComputerSystem
    if ($computerSystem -eq $null) {
        Write-Verbose "$ComputerName : Could not connect to WMI (Win32_ComputerSystem)"
        $result.WmiStatus = $error[0].Exception
        return $result
    } else {
        $result.WmiStatus = "Success"
        $result.Manufacturer = $computerSystem.Manufacturer
        $result.Model = $computerSystem.Model
    }

    $error.Clear()
    $operatingSystem = Get-WmiObject -ComputerName $ComputerName -Class Win32_OperatingSystem
    if ($operatingSystem -eq $null) {
        Write-Verbose "$ComputerName : Could not connect to WMI (Win32_OperatingSystem)"
        $result.WmiStatus = $error[0].Exception
        return $result
    } else {
        $result.WmiStatus = "Success"
        $result.OSVersion = $operatingSystem.Version
        $result.OSName = $operatingSystem.Name.Split("|")[0]
    }

    $error.Clear()
    $serverName = Invoke-Command -ComputerName $ComputerName `
                    -ScriptBlock { Get-Item Env:ComputerName } `
                    -ErrorAction SilentlyContinue
    if ($serverName -eq $null) {
        Write-Verbose "$ComputerName : Could not connect to server via Remote PowerShell"
        $result.PSRemotingStatus = "Failed"
        $result.PSRemotingError = $error[0].Exception
    } else {
        $result.PSRemotingStatus = "Success"
    }

    return $result
}
