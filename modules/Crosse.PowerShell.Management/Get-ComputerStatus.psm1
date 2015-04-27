function Get-ComputerStatus {
    [CmdletBinding()]
    param (
            [Parameter(Mandatory=$true,
                ValueFromPipeline=$true)]
            [string]
            $ComputerName,

            [Parameter(Mandatory=$false)]
            [switch]
            # Indicates whether to force all tests to run.
            $Force=$false,

            [switch]
            # Test extraneous services such as Remote Registry, Remote Firewall, Remote Services, file sharing, etc.
            $TestServices = $true
        )

    BEGIN {
        $ping = New-Object System.Net.NetworkInformation.Ping
    }

    PROCESS {
        $result = New-Object PSObject -Property @{
            "DnsStatus"                         = "NotFound"
            "HostName"                          = $ComputerName
            "IpAddress"                         = $null
            "PingStatus"                        = $null
            "PingRoundTripTime"                 = $null
            "WmiStatus"                         = $null
            "WmiError"                          = $null
            "Manufacturer"                      = $null
            "Model"                             = $null
            "OSVersion"                         = $null
            "OSName"                            = $null
            "PSRemotingStatus"                  = $null
            "PSRemotingError"                   = $null
            "WinRSStatus"                       = $null
            "WinRSError"                        = $null
            # Services to test
            "FileSharingEnabled"                = $false
            "RemoteDesktopEnabled"              = $false
        }

        Write-Verbose "$ComputerName : starting."

        try {
            Write-Verbose "$ComputerName : Resolving DNS name"
            $dns = [System.Net.Dns]::GetHostEntry($ComputerName)
            $result.DnsStatus = "Found"
            $result.HostName = $dns.HostName
            $result.IpAddress = $dns.AddressList[0].IpAddressToString
        } catch {
            $result.DnsStatus = "NotFound"
            Write-Warning "$ComputerName : Name not found in DNS."
            if (!$Force) {
                return $result
            }
        }

        $pingResult = $ping.Send($ComputerName, 1000)
        $result.PingStatus = $pingResult.Status
        $result.PingRoundTripTime = $pingResult.RoundTripTime

        if ($pingResult.Status -ne 'Success') {
            Write-Warning "$ComputerName : Did not respond to ping."
            if (!$Force) {
                return $result
            }
        }

        Write-Verbose "$ComputerName : Attempting WMI connections"
        try {
            $computerSystem = Get-WmiObject -ComputerName $ComputerName -Class Win32_ComputerSystem -ErrorAction Stop

            $result.WmiStatus = "Success (ComputerSystem)"
            $result.Manufacturer = $computerSystem.Manufacturer
            $result.Model = $computerSystem.Model

            try {
                $operatingSystem = Get-WmiObject -ComputerName $ComputerName -Class Win32_OperatingSystem -ErrorAction Stop

                $result.WmiStatus = "Success"
                $result.OSVersion = $operatingSystem.Version
                $result.OSName = $operatingSystem.Name.Split("|")[0]
            } catch {
                Write-Warning "$ComputerName : Could not connect to WMI (Win32_OperatingSystem)"
                $result.WmiStatus += "/Failed (OperatingSystem)"
                $result.WmiError += $_
                if (!$Force) {
                    return $result
                }
            }
        } catch {
            Write-Warning "$ComputerName : Could not connect to WMI (Win32_ComputerSystem)"
            $result.WmiStatus = "Failed"
            $result.WmiError = $_
            if (!$Force) {
                return $result
            }
        }

        Write-Verbose "$ComputerName : Attempting Remote PowerShell (WS-Man) connection"
        try {
            $null = Test-WSMan $ComputerName -ErrorAction Stop
            $result.PSRemotingStatus = "Success"
        } catch {
            Write-Warning "$ComputerName : Could not connect to server via remote PowerShell"
            $result.PSRemotingStatus = "Failed"
            $result.PSRemotingError = ([xml]$_.Exception.Message).WSManFault.Message
        }

        Write-Verbose "$ComputerName : Attempting WinRS connection"
        $error.Clear()
        $cmdResult = winrs.exe -r:$ComputerName "SET COMPUTERNAME"
        if ($? -and ![String]::IsNullOrEmpty($cmdResult)) {
            $result.WinRSStatus = "Success"
        } else {
            Write-Warning "$ComputerName : Could not connect to server via WinRS"
            $result.WinRSStatus = "Failed"
            $result.WinRSError = $error[0].Exception
        }

        if ($TestServices) {
            # FileSharingEnabled
            Write-Verbose "$ComputerName : Testing file share access"
            if (Test-Path "\\$ComputerName\ADMIN$\") {
                $result.FileSharingEnabled = $true
            } else {
                Write-Warning "$ComputerName : Could not enumerate ADMIN$ file share"
            }

            # RemoteDesktopEnabled
            Write-Verbose "$ComputerName : Testing Remote Desktop"
            if (Test-NetConnection -ComputerName $ComputerName -CommonTCPPort RDP) {
                $result.RemoteDesktopEnabled = $true
            } else {
                Write-Warning "$ComputerName : Could not connect to RDP port"
            }
        }

        return $result | Select @SelectHash
    }
}
