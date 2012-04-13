function Get-SqlServerProperties {
    param (
            [Parameter(Mandatory=$true)]
            [ValidateNotNull()]
            [string]
            $SqlServerInstance
          )
    $server, $instance = $SqlServerInstance.Split('\')

    Write-Verbose "Server: $server"

    if ([String]::IsNullOrEmpty($instance)) {
        Write-Verbose "Instance: (Default Instance)"
        $serviceName = 'MSSQLSERVER'
    } else {
        Write-Verbose "Instance: $instance"
        $serviceName = 'MSSQL${0}' -f $instance
    }

    $service = Get-WmiObject Win32_Service -ComputerName $server -Filter "Name = '$serviceName'"
    if ($service -eq $null) {
        throw "Could not find the SQL Server service on $server for the specified instance"
    } else {
        $serviceAccount = $service.StartName
        Write-Verbose "Service $($service.Name) on $server is running as $serviceAccount"
    }

    return New-Object PSObject -Property @{
            Server          = $server
            Instance        = $instance
            ServiceName     = $serviceName
            ServiceAccount  = $serviceAccount
        }
}

function Get-SqlDatabaseProperties {
    param (
            [ValidateNotNull()]
            [System.Data.SqlClient.SqlConnection]
            $SqlConnection,

            [ValidateNotNull()]
            [Object]
            $Database
          )
    $cmd = "SELECT compatibility_level, collation_name, state_desc, recovery_model, recovery_model_desc FROM sys.databases WHERE name = '$Database'"
    $props = @(Send-SqlQuery $SqlConnection $cmd)
    
    if ($props.Count -gt 1) {
        throw "Multiple results returned when getting properties for database $Database"
    }
    return $props
}

function Open-SqlConnection {
    param (
            [Parameter(Mandatory=$true)]
            [string]
            $SqlServerInstance
          )

    $connString = "Data Source={0};Initial Catalog=master;Integrated Security=SSPI;" -f $SqlServerInstance
    $conn = New-Object System.Data.SqlClient.SqlConnection $connString
    Write-Verbose "Opening connection to $SqlServerInstance"
    $conn.Open()
    if ($conn.State -ne 'Open') {
        throw "Could not open SQL connection to the server $SqlServerInstance"
    }
    return $conn
}

function Close-SqlConnection {
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
    param (
            [ValidateNotNull()]
            [System.Data.SqlClient.SqlConnection]
            $SqlConnection,

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
    param (
            [ValidateNotNull()]
            [System.Data.SqlClient.SqlConnection]
            $SqlConnection,

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
    param (
            [ValidateNotNull()]
            [System.Data.SqlClient.SqlConnection]
            $SqlConnection,

            [ValidateNotNullOrEmpty()]
            [string]
            $Command
          )

    $cmd = $SqlConnection.CreateCommand()
    $cmd.CommandText = $Command

    #Write-Verbose "Executing: `"$Command`""

    try {
        $reader = $cmd.ExecuteReader()
        $rows = @()
        while ($reader.Read()) {
            $row = New-Object Object[] $reader.FieldCount
            $reader.GetValues($row) | Out-Null

            $values = @{}
            for ($i = 0; $i -lt $reader.FieldCount; $i++) {
                $values[$reader.GetName($i)] = $row[$i]
            }

            $rows += New-Object PSObject -Property $values
        }
    } catch {
        throw
    } finally {
        $cmd.Dispose()
        $reader.Close()
    }

    #Write-Verbose "Retrieved $($rows.Count) row(s) from the database"
    return $rows
}

