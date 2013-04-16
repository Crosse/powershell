function Get-RegistryKeyValue {
    [CmdletBinding()]
    param (
            [Parameter(Mandatory=$false,
                ValueFromPipeline=$true,
                ParameterSetName="WMI")]
            [Parameter(Mandatory=$false,
                ValueFromPipeline=$true,
                ParameterSetName="RemoteRegistry")]
            [Parameter(Mandatory=$false,
                ValueFromPipeline=$true,
                ParameterSetName="UseBestMethod")]
            [string]
            $ComputerName,

            [Parameter(Mandatory=$true,
                ParameterSetName="WMI")]
            [switch]
            $UseWmi = $false,

            [Parameter(Mandatory=$true,
                ParameterSetName="RemoteRegistry")]
            [switch]
            $UseRemoteRegistry = $false,

            [Parameter(Mandatory=$true)]
            [Microsoft.Win32.RegistryHive]
            $Hive,

            [Parameter(Mandatory=$true)]
            [ValidateNotNullOrEmpty()]
            [string]
            $Key,

            [Parameter(Mandatory=$true)]
            [ValidateNotNullOrEmpty()]
            [string]
            $ValueName,

            [Parameter(Mandatory=$true,
                ParameterSetName="UseBestMethod")]
            [Parameter(Mandatory=$true,
                ParameterSetName="WMI")]
            [ValidateSet("Binary", "Dword", "ExpandedString",
                         "MultiString", "Qword", "String")]
            [string]
            $ValueType
          )

    if (!$UseRemoteRegistry -and !$UseWmi) {
        $UseRemoteRegistry = $true
    }

    if ($UseRemoteRegistry) {
        if ([String]::IsNullOrEmpty($ComputerName)) {
            $ComputerName = [System.Net.Dns]::GetHostName()
            Write-Verbose "Using Registry service"
            $reg = [Microsoft.Win32.RegistryKey]::OpenBaseKey($Hive, "Default")
        } else {
            Write-Verbose "Using Remote Registry service"
            $reg = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey($Hive, $ComputerName)
        }

        if ($reg -eq $null) {
            Write-Error "Could not open base key"
            return
        }
        $subkey = $reg.OpenSubKey($Key)
        if ($subkey -eq $null) {
            Write-Error "Could not find subkey"
            return
        }
        if ($ValueName -notin $subkey.GetValueNames()) {
            Write-Error "Value not found in subkey"
            return
        }
        $result = $subkey.GetValue($ValueName)
        $reg.Close()
    }
    
    if ($result -eq $null -or $UseWmi) {
        if ([String]::IsNullOrEmpty($ComputerName)) {
            $ComputerName = [System.Net.Dns]::GetHostName()
        }

        Write-Verbose "Using WMI"
        [UInt32]$hkey = 0x7FFFFFFF + 1

        switch ($Hive) {
            "ClassesRoot"       { $hkey += 0 }
            "CurrentUser"       { $hkey += 1 }
            "LocalMachine"      { $hkey += 2 }
            "Users"             { $hkey += 3 }
            "PerformanceData"   { $hkey += 4 }
            "CurrentConfig"     { $hkey += 5 }
            "DynData"           { $hkey += 6 }
        }

        $reg = Get-WmiObject -List -Namespace "ROOT\default" -Class StdRegProv -ComputerName $ComputerName -ErrorAction Stop
        if ($reg -eq $null) {
            Write-Error "Could not connect to remote registry via WMI"
            return
        }

        switch ($ValueType) {
            "Binary"            { $result = $reg.GetBinaryValue($hkey, $Key, $ValueName) }
            "Dword"             { $result = $reg.GetDWORDValue($hkey, $Key, $ValueName) }
            "ExpandedString"    { $result = $reg.GetExpandedStringValue($hkey, $Key, $ValueName) }
            "MultiString"       { $result = $reg.GetMultiStringValue($hkey, $Key, $ValueName) }
            "Qword"             { $result = $reg.GetQWORDValue($hkey, $Key, $ValueName) }
            "String"            { $result = $reg.GetStringValue($hkey, $Key, $ValueName) }
        }

        if ($result -eq $null) {
            Write-Error "Value not found in subkey"
            return
        }

        $result = $result.uValue
    }

    return New-Object -TypeName PSObject -Property @{
        ComputerName = $ComputerName
        $ValueName = $result
    }
}
