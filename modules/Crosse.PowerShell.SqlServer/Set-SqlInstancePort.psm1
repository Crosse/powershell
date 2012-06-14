function Set-SqlInstancePort {
    [CmdletBinding()]
    param (
            [Parameter(Mandatory=$true)]
            [string]
            $Name,

            [Parameter(Mandatory=$true)]
            [ValidateRange(1,65535)]
            [int]
            $StaticPort
          )

    $Name = $Name.ToUpper()

    $baseKey = "HKLM:\SOFTWARE\Microsoft\Microsoft SQL Server\"
    if ((Test-Path $baseKey) -eq $false) {
        throw "Cannot access base SQL Server registry key.  SQL Server may not be installed."
    }

    $lName = (Get-ItemProperty -Path (Join-Path $baseKey "Services\SQL Server") -Name LName).LName
    if ($lName -eq $null) {
        throw "Cannot query for service name prefix."
    }

    $instanceNames = Join-Path $baseKey "Instance Names\SQL"
    $instanceName = (Get-ItemProperty -Path $instanceNames).${Name}
    if ($instanceName -eq $null) {
        throw "Could not find instance name $Name"
    }

    $key = Resolve-Path "HKLM:\SOFTWARE\Microsoft\Microsoft SQL Server\${instanceName}\MSSQLServer\SuperSocketNetLib\Tcp"
    if ($key -eq $null) {
        throw "Could not find registry key for instance $Name"
    }

    foreach ($item in (Get-ChildItem $key)) {
        $ipAddress = (Get-ItemProperty -Path $item.PSPath).IpAddress
        Write-Verbose "Setting static port to $StaticPort for $ipAddress"
        Set-ItemProperty -LiteralPath $item.PSPath -Name TcpPort -Value $StaticPort
        Write-Verbose "Disabling dynamic ports for $ipAddress"
        Set-ItemProperty -LiteralPath $item.PSPath -Name TcpDynamicPorts -Value $null
    }

    $service = Get-Service "${lName}${Name}"
    Write-Warning "You must restart the '$($service.Name)' service for the changes to take effect."
}

