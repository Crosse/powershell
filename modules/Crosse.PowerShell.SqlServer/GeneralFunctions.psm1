function Get-SqlServerProperties {
    [CmdletBinding()]
    param (
            [Parameter(Mandatory=$true)]
            [ValidateNotNull()]
            [string]
            $Server,

            [Parameter(Mandatory=$false)]
            [ValidateNotNull()]
            [System.Management.Automation.PSCredential]
            $Credential
          )
    $serverName, $instance = $Server.Split('\')

    Write-Verbose "Server: $serverName"

    if ([String]::IsNullOrEmpty($instance)) {
        Write-Verbose "Instance: (Default Instance)"
        $serviceName = 'MSSQLSERVER'
    } else {
        Write-Verbose "Instance: $instance"
        $serviceName = 'MSSQL${0}' -f $instance
    }

    if ($Credential) {
        $service = Get-WmiObject Win32_Service -ComputerName $serverName Credential $Credential -Filter "Name = '$serviceName'"
    } else {
        $service = Get-WmiObject Win32_Service -ComputerName $serverName -Filter "Name = '$serviceName'"
    }

    if ($service -eq $null) {
        throw "Could not find the SQL Server service on $serverName for the specified instance"
    } else {
        $serviceAccount = $service.StartName
        Write-Verbose "Service $($service.Name) on $serverName is running as $serviceAccount"
    }

    return New-Object PSObject -Property @{
            Server          = $serverName
            Instance        = $instance
            ServiceName     = $serviceName
            ServiceAccount  = $serviceAccount
        }
}

function Get-SqlDatabaseProperties {
    [CmdletBinding()]
    param (
            [Parameter(Mandatory=$true)]
            [ValidateNotNull()]
            [System.Data.SqlClient.SqlConnection]
            $SqlConnection,

            [Parameter(Mandatory=$false)]
            [ValidateNotNullOrEmpty()]
            [string]
            $Database
          )

    if ([String]::IsNullOrEmpty($Database)) {
        $Database = $SqlConnection.Database
    }

    $cmd = "SELECT * FROM sys.databases WHERE name = '$Database'"
    $props = @(Send-SqlQuery $SqlConnection $cmd)
    
    if ($props.Count -gt 1) {
        throw "Multiple results returned when getting properties for database $Database"
    }
    return $props
}

function Open-SqlConnection {
    [CmdletBinding()]
    param (
            [Parameter(Mandatory=$true)]
            [ValidateNotNullOrEmpty()]
            [string]
            $Server,

            [Parameter(Mandatory=$false)]
            [ValidateNotNullOrEmpty()]
            [string]
            $Database="master",

            [Parameter(Mandatory=$false)]
            [switch]
            $Async
          )

    $connString = "Data Source={0};Initial Catalog={1};Integrated Security=SSPI;" -f $Server, $Database
    if ($Async) {
        $connString += "Asynchronous Processing=true;"
    }
    $conn = New-Object System.Data.SqlClient.SqlConnection $connString
    Write-Verbose "Opening connection to $Server"
    $conn.Open()
    if ($conn.State -ne 'Open') {
        throw "Could not open SQL connection to the server $Server"
    }
    return $conn
}

function Close-SqlConnection {
    [CmdletBinding()]
    param (
            [Parameter(Mandatory=$false)]
            [System.Data.SqlClient.SqlConnection]
            $SqlConnection
          )

    if ($SqlConnection) {
        if ($SqlConnection.State -eq 'Open') {
            Write-Verbose "Closing connection to $($SqlConnection.Datasource)"
            $SqlConnection.Close()
        }
        $SqlConnection.Dispose()
    }
}

function Send-SqlScalarQuery {
    [CmdletBinding()]
    param (
            [Parameter(Mandatory=$true)]
            [ValidateNotNull()]
            [System.Data.SqlClient.SqlConnection]
            $SqlConnection,

            [Parameter(Mandatory=$true)]
            [ValidateNotNullOrEmpty()]
            [string]
            $Command
          )
    
    $cmd = $SqlConnection.CreateCommand()
    $cmd.CommandText = $Command

    #Write-Verbose "Executing: `"$Command`""
    try {
        $result = $cmd.ExecuteScalar()
    } catch {
        throw
    } finally {
        $cmd.Dispose()
    }

    return $result
}

function Send-SqlNonQuery {
    [CmdletBinding()]
    param (
            [Parameter(Mandatory=$true)]
            [ValidateNotNull()]
            [System.Data.SqlClient.SqlConnection]
            $SqlConnection,

            [Parameter(Mandatory=$true)]
            [ValidateNotNullOrEmpty()]
            [string]
            $Command
          )
    
    $cmd = $SqlConnection.CreateCommand()
    $cmd.CommandText = $Command

    #Write-Verbose "Executing: `"$Command`""
    try {
        $result = $cmd.ExecuteNonQuery()
    } catch {
        throw
    } finally {
        $cmd.Dispose()
    }

    return $result
}

function Send-SqlQuery {
    [CmdletBinding()]
    param (
            [Parameter(Mandatory=$true)]
            [ValidateNotNull()]
            [System.Data.SqlClient.SqlConnection]
            $SqlConnection,

            [Parameter(Mandatory=$true)]
            [ValidateNotNullOrEmpty()]
            [string]
            $Command
          )

    $cmd = $SqlConnection.CreateCommand()
    $cmd.CommandText = $Command

    #Write-Verbose "Executing: `"$Command`""

    $reader = $null
    try {
        $reader = $cmd.ExecuteReader()
        $rows = @()
        while ($reader.Read()) {
            $row = New-Object Object[] $reader.FieldCount
            $reader.GetValues($row) | Out-Null

            $values = @{}
            for ($i = 0; $i -lt $reader.FieldCount; $i++) {
                if ([String]::IsNullOrEmpty($reader.GetName($i))) {
                    $colName = "Column_" + ($i + 1)
                } else {
                    $colName = $reader.GetName($i)
                }

                if ($row[$i] -is [DBNull]) {
                    $values[$colName] = $null
                } else {
                    $values[$colName] = $row[$i]
                }
            }

            $rows += New-Object PSObject -Property $values
        }
    } catch {
        throw
    } finally {
        $cmd.Dispose()
        if ($reader -ne $null) {
            $reader.Close()
        }
    }

    #Write-Verbose "Retrieved $($rows.Count) row(s) from the database"
    return $rows
}

