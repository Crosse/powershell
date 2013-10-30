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

function Get-SqlServerInstance {
    [CmdletBinding(DefaultParameterSetName="NamedInstance")]
    param (
            [Parameter(Mandatory=$true,
                ValueFromPipeline=$true)]
            [ValidateNotNullOrEmpty()]
            [string]
            $ComputerName,

            [Parameter(Mandatory=$false,
                ValueFromPipelineByPropertyName=$true,
                ParameterSetName="NamedInstance")]
            [string]
            $InstanceName,

            [Parameter(Mandatory=$false,
                ParameterSetName="DefaultInstance")]
            [switch]
            $DefaultInstance
          )
    PROCESS {
        if ([String]::IsNullOrEmpty($InstanceName)) {
            if ($DefaultInstance) {
                Write-Verbose "Querying for the default instance of SQL Server on $ComputerName"
                $services = Get-Service -ComputerName $ComputerName -Name "MSSQLSERVER"
            } else {
                Write-Verbose "Querying list of all SQL Server instances on $ComputerName"
                $services = Get-Service -ComputerName $ComputerName -DisplayName "SQL Server (*"
            }
        } else {
            Write-Verbose "Querying for the named instance $InstanceName on $ComputerName"
            $services = Get-Service -ComputerName $ComputerName -DisplayName "SQL Server ($InstanceName)*"
        }
        foreach ($service in $services) {
            $info = New-Object PSObject -Property @{
                InstanceName        = $null
                ProductVersion      = $null
                ProductLevel        = $null
                Edition             = $null
                IsClustered         = $null
                IsAlwaysOnEnabled   = $null
                IsWindowsAuthOnly   = $null
            }

            if ($service.Name.Contains("$")) {
                $instance = $ComputerName + "\" + $service.Name.Replace("MSSQL$", "")
            } else {
                $instance = $ComputerName
            }

            if ($service.Status -ne "Running") {
                Write-Warning "SQL Server $($service.Name) is in the $($service.Status)` state."
                Write-Output $info
                continue
            }

            try {
                Write-Verbose "Getting version information for instance $instance"
                $conn = Open-SqlConnection -Server $instance
                $version = Send-SqlQuery -SqlConnection $conn -Command `
                "SELECT @@SERVERNAME AS ServerName
                        , SERVERPROPERTY('ProductVersion') AS ProductVersion
                        , SERVERPROPERTY('ProductLevel') AS ProductLevel
                        , SERVERPROPERTY('Edition') AS Edition
                        , SERVERPROPERTY('IsClustered') AS IsClustered
                        , SERVERPROPERTY('IsHadrEnabled') AS IsAlwaysOnEnabled
                        , SERVERPROPERTY('IsIntegratedSecurityOnly') AS IsWindowsAuthOnly
                "

                $info.InstanceName        = $version.ServerName
                $info.ProductVersion      = $version.ProductVersion
                $info.ProductLevel        = $version.ProductLevel
                $info.Edition             = $version.Edition
                $info.IsClustered         = [Nullable[Boolean]]$version.IsClustered
                $info.IsAlwaysOnEnabled   = [Nullable[Boolean]]$version.IsAlwaysOnEnabled
                $info.IsWindowsAuthOnly   = [Nullable[Boolean]]$version.IsWindowsAuthOnly

                Write-Output $info
            } catch {
                throw
            } finally {
                Close-SqlConnection -SqlConnection $conn
            }
        }
    }
}

function Get-SqlDatabase {
    [CmdletBinding()]
    param (
            [Parameter(Mandatory=$true,
                ValueFromPipeline=$true,
                ValueFromPipelineByPropertyName=$true)]
            [ValidateNotNullOrEmpty()]
            [string]
            $InstanceName,

            [Parameter(Mandatory=$false,
                ValueFromPipelineByPropertyName=$true)]
            [string]
            $Database,

            [Parameter(Mandatory=$false)]
            [switch]
            $OnlyUserDatabases
          )

    PROCESS {
        try {
            $conn = Open-SqlConnection -Server $InstanceName
            $query = "SELECT @@SERVERNAME AS InstanceName
                        , name
                        , database_id
                        , create_date
                        , compatibility_level
                        , state_desc
                        , recovery_model_desc
                        , page_verify_option_desc
                      FROM sys.databases"

            $where = @()
            if (![String]::IsNullOrEmpty($Database)) {
                $where += "name = '$Database'"
            }

            if ($OnlyUserDatabases) {
                $where += "LEN(owner_sid) > 1"
            }

            if ($where.Count -gt 0) {
                $whereClause = "WHERE " + ($where -join " AND ")
                $query += "
                      $whereClause"
            }

            Write-Verbose $query

            $databases = Send-SqlQuery -SqlConnection $conn -Command $query
            foreach ($db in $databases) {
                New-Object PSObject -Property @{
                    InstanceName        = $db.InstanceName
                    Name                = $db.name
                    DatabaseId          = $db.database_id
                    WhenCreated         = [DateTime]::SpecifyKind([DateTime]$db.create_date, 'Utc').ToLocalTime()
                    CompatibilityLevel  = $db.compatibility_level
                    State               = $db.state_desc
                    RecoveryModel       = $db.recovery_model_desc
                    PageVerifyOption    = $db.page_verify_option_desc
                }
            }
        } catch {
            throw
        } finally {
            Close-SqlConnection -SqlConnection $conn
        }
    }
}
