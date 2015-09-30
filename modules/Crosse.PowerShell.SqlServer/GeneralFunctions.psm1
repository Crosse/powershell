################################################################################
#
# Copyright (c) 2013 Seth Wright <wrightst@jmu.edu>
#
# Permission to use, copy, modify, and distribute this software for any
# purpose with or without fee is hereby granted, provided that the above
# copyright notice and this permission notice appear in all copies.
#
# THE SOFTWARE IS PROVIDED "AS IS" AND THE AUTHOR DISCLAIMS ALL WARRANTIES
# WITH REGARD TO THIS SOFTWARE INCLUDING ALL IMPLIED WARRANTIES OF
# MERCHANTABILITY AND FITNESS. IN NO EVENT SHALL THE AUTHOR BE LIABLE FOR
# ANY SPECIAL, DIRECT, INDIRECT, OR CONSEQUENTIAL DAMAGES OR ANY DAMAGES
# WHATSOEVER RESULTING FROM LOSS OF USE, DATA OR PROFITS, WHETHER IN AN
# ACTION OF CONTRACT, NEGLIGENCE OR OTHER TORTIOUS ACTION, ARISING OUT OF
# OR IN CONNECTION WITH THE USE OR PERFORMANCE OF THIS SOFTWARE.
#
################################################################################

function Get-SqlServerProperties {
    [CmdletBinding()]
    param (
            [Parameter(Mandatory=$true)]
            [ValidateNotNull()]
            [Alias("Server")]
            [string]
            $InstanceName,

            [Parameter(Mandatory=$false)]
            [ValidateNotNull()]
            [System.Management.Automation.PSCredential]
            $Credential
          )
    $serverName, $instance = $InstanceName.Split('\')

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
    $props = @(Send-SqlQuery -SqlConnection $SqlConnection -Query $cmd)
    
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
            [Alias("Server")]
            [string]
            $InstanceName,

            [Parameter(Mandatory=$false)]
            [ValidateNotNullOrEmpty()]
            [string]
            $Database = "master",

            [Parameter(Mandatory=$false)]
            [switch]
            $Async
          )

    $connString = "Data Source={0};Initial Catalog={1};Integrated Security=SSPI;" -f $InstanceName, $Database
    if ($Async) {
        $connString += "Asynchronous Processing=true;"
    }
    $conn = New-Object System.Data.SqlClient.SqlConnection $connString
    Write-Verbose "Opening connection to $InstanceName"
    $conn.Open()
    if ($conn.State -ne 'Open') {
        throw "Could not open SQL connection to the server $InstanceName"
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

################################################################################
<#
    .SYNOPSIS
    Sends a SQL statement to a database and returns the result.

    .DESCRIPTION
    Sends a SQL query to a SQL Server database and returns the result.  Useful
    for statements that do not return row data, such as UPDATEs and DELETEs.

    .INPUTS
    System.Data.SqlClient.SqlConnection.  An open SQL connection to run the query against.

    System.String.  The query to run.

    .OUTPUTS
    A System.String of the response from the server, if any.
#>
################################################################################
function Send-SqlNonQuery {
    [CmdletBinding(DefaultParameterSetName="Implicit")]
    param (
            [Parameter(Mandatory=$true,
                ParameterSetName="Implicit")]
            [ValidateNotNull()]
            [string]
            # The MSSQL instance to run the statement against.  This can be
            # either a named instance ("SERVER\INSTANCENAME") or a default
            # instance ("SERVER").
            $InstanceName,

            [Parameter(Mandatory=$false,
                ParameterSetName="Implicit")]
            [ValidateNotNull()]
            [string]
            # The database to run the statement against.  If not specified,
            # defaults to the "master" database.
            $Database = "master",

            [Parameter(Mandatory=$true,
                ParameterSetName="Explicit")]
            [ValidateNotNull()]
            [System.Data.SqlClient.SqlConnection]
            # A connection to an MSSQL server that has been previously opened
            # with the Open-SqlConnection cmdlet.
            $SqlConnection,

            [Parameter(Mandatory=$true)]
            [ValidateNotNullOrEmpty()]
            [Alias("Command")]
            [Alias("Query")]
            [string]
            # The SQL query to run.
            $Statement
          )
    
    if ($PSCmdlet.ParameterSetName -eq 'Implicit') {
        try {
            $conn = Open-SqlConnection -Server $InstanceName -Database $Database
        } catch {
            if ($conn -ne $null) {
                Close-SqlConnection $conn
            }
            throw
        }
    } else {
        $conn = $SqlConnection
    }

    $cmd = $conn.CreateCommand()
    $cmd.CommandText = $Statement

    try {
        Write-Verbose "Executing: `"$Statement`""
        $result = $cmd.ExecuteNonQuery()
    } catch {
        throw
    } finally {
        $cmd.Dispose()
    }

    if ($PSCmdlet.ParameterSetName -eq "ImplicitQuery") {
        Close-SqlConnection $conn
    }

    return $result
}

################################################################################
<#
    .SYNOPSIS
    Sends a SQL query to a database and returns the result.

    .DESCRIPTION
    Sends a SQL query to a SQL Server database and returns the result.

    .INPUTS
    System.Data.SqlClient.SqlConnection.  An open SQL connection to run the query against.

    System.String.  The query to run.

    .OUTPUTS
    An array of PSObjects representing the returned data.
#>
################################################################################
function Send-SqlQuery {
    [CmdletBinding(DefaultParameterSetName="ImplicitQuery")]
    param (
            [Parameter(Mandatory=$true,
                ParameterSetName="ImplicitQuery")]
            [ValidateNotNull()]
            [string]
            # The MSSQL instance to run the query against.  This can be either
            # a named instance ("SERVER\INSTANCENAME") or a default instance
            # ("SERVER").
            $InstanceName,

            [Parameter(Mandatory=$false,
                ParameterSetName="ImplicitQuery")]
            [ValidateNotNull()]
            [string]
            # The database to run the query against.  If not specified, defaults
            # to the "master" database.
            $Database = "master",

            [Parameter(Mandatory=$true,
                ParameterSetName="ExplicitQuery")]
            [ValidateNotNull()]
            [System.Data.SqlClient.SqlConnection]
            # A connection to an MSSQL server that has been previously opened
            # with the Open-SqlConnection cmdlet.
            $SqlConnection,

            [Parameter(Mandatory=$false)]
            [switch]
            # Return only the left-most column of the first row of data.  Also
            # called a "scalar" query.
            $SingleResult,

            [Parameter(Mandatory=$true)]
            [ValidateNotNullOrEmpty()]
            [Alias("Command")]
            [string]
            # The SQL query to run.
            $Query
          )

    if ($PSCmdlet.ParameterSetName -eq 'ImplicitQuery') {
        try {
            $conn = Open-SqlConnection -Server $InstanceName -Database $Database
        } catch {
            if ($conn -ne $null) {
                Close-SqlConnection $conn
            }
            throw
        }
    } else {
        $conn = $SqlConnection
    }

    $cmd = $conn.CreateCommand()
    $cmd.CommandText = $Query

    $reader = $null
    try {
        if ($SingleResult) {
            $result = $cmd.ExecuteScalar()
        } else {
            Write-Verbose "Executing: `"$Query`" against $($conn.Datasource)"
            $reader = $cmd.ExecuteReader()
            $result = @()
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
                $result += New-Object PSObject -Property $values
            }
        }
    } catch {
        throw
    } finally {
        if (!$SingleResult -and $reader) {
            $reader.Close()
            $reader.Dispose()
        }
        $cmd.Dispose()
    }

    if ($PSCmdlet.ParameterSetName -eq "ImplicitQuery") {
        Close-SqlConnection $conn
    }

    return $result
}

function Get-SqlServerInstance {
    [CmdletBinding(DefaultParameterSetName="NamedInstance")]
    param (
            [Parameter(Mandatory=$false,
                ValueFromPipeline=$true)]
            [ValidateNotNullOrEmpty()]
            [string]
            $ComputerName = "localhost",

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

            [Parameter(Mandatory=$false,
                ParameterSetName="User")]
            [switch]
            $OnlyUserDatabases,

            [Parameter(Mandatory=$false,
                ParameterSetName="System")]
            [switch]
            $OnlySystemDatabases
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

            if ($OnlySystemDatabases) {
                $where += "LEN(owner_sid) = 1"
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
