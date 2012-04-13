function New-SqlServerMirroringSession {
[CmdletBinding()]
    param (
            [Parameter(Mandatory=$true)]
            [string]
            $PrimaryServerInstance,

            [Parameter(Mandatory=$true)]
            [string]
            $MirrorServerInstance,

            [Parameter(Mandatory=$false)]
            [string]
            $WitnessServerInstance,

            [Parameter(Mandatory=$true)]
            [string]
            $Database,

            [Parameter(Mandatory=$true)]
            [ValidateRange(1025, 65536)]
            [int]
            $EndpointPort
            )

    BEGIN {
        $primaryServerProperties = Get-SqlServerProperties $PrimaryServerInstance
        $mirrorServerProperties = Get-SqlServerProperties $MirrorServerInstance
        if ($WitnessServerInstance) {
            $witnessServerProperties = Get-SqlServerProperties $WitnessServerInstance
        }

        try {
            $primaryConn = Open-SqlConnection $PrimaryServerInstance
            $mirrorConn = Open-SqlConnection $MirrorServerInstance

            if ($WitnessServerInstance) {
                $witnessConn = Open-SqlConnection $WitnessServerInstance
            }
        } catch {
            $err = $_
            Close-SqlConnection $primaryConn
            Close-SqlConnection $mirrorConn
            Close-SqlConnection $witnessConn
            throw $_
        }
    }

    PROCESS {
        $primaryDb = Get-SqlDatabaseProperties $primaryConn $Database
        $mirrorDb = Get-SqlDatabaseProperties $mirrorConn $Database

        if ($primaryDb -eq $null) {
            Write-Error "Could not find database $Database on primary instance $PrimaryServerInstance"
            return
        }
        if ($mirrorDb -eq $null) {
            Write-Error "Could not find database $Database on mirror instance $MirrorServerInstance"
            return
        }

        Write-Verbose "Primary database collation: $($primaryDb.collation_name)"
        Write-Verbose "Primary database recovery model: $($primaryDb.recovery_model_desc)"

        Write-Verbose "Mirror database collation: $($mirrorDb.collation_name)"
        Write-Verbose "Mirror database recovery model: $($mirrorDb.recovery_model_desc)"

        $retval = Confirm-DatabaseMirroringReadiness $primaryDb $mirrorDb
        if ($retval -eq $false) {
            return
        } else {
            Write-Verbose "Primary and mirror databases passed validation."
        }

        try {
            $primaryEndpoint = New-MirroringEndpoint $primaryConn $EndpointPort
            Write-Verbose "Using endpoint `"$($primaryEndpoint.name)`", port $($primaryEndpoint.port) for $($primaryConn.Datasource)"
            Grant-PrivilegesOnEndpoint $primaryConn $mirrorServerProperties.ServiceAccount $primaryEndpoint

            $mirrorEndpoint = New-MirroringEndpoint $mirrorConn $EndpointPort
            Write-Verbose "Using endpoint `"$($mirrorEndpoint.name)`", port $($mirrorEndpoint.port) for $($mirrorConn.Datasource)"
            Grant-PrivilegesOnEndpoint $mirrorConn $primaryServerProperties.ServiceAccount $mirrorEndpoint

            if ($WitnessServerInstance) {
                $witnessEndpoint = New-MirroringEndpoint $witnessConn $EndpointPort
                Write-Verbose "Using endpoint `"$($witnessEndpoint.name)`", port $($witnessEndpoint.port) for $($witnessConn.Datasource)"

                # Grant privs on the primary instance to the witness instance.
                Grant-PrivilegesOnEndpoint $primaryConn $witnessServerProperties.ServiceAccount $primaryEndpoint
                # Grant privs on the witness instance to the primary instance.
                Grant-PrivilegesOnEndpoint $witnessConn $primaryServerProperties.ServiceAccount $witnessEndpoint

                # Grant privs on the mirror instance to the witness instance.
                Grant-PrivilegesOnEndpoint $mirrorConn $witnessServerProperties.ServiceAccount $mirrorEndpoint
                # Grant privs on the witness instance to the primary instance.
                Grant-PrivilegesOnEndpoint $witnessConn $mirrorServerProperties.ServiceAccount $witnessEndpoint
            }
        } catch {
            Write-Error $_
            return
        }

        try {
            Write-Verbose "Starting mirroring session on mirror server"
            $setPartnerMirror = "ALTER DATABASE $Database SET PARTNER = 'TCP://$($PrimaryServerInstance.Split('\')[0]):$($primaryEndpoint.port)'"
            Send-SqlNonQuery $mirrorConn $setPartnerMirror | Out-Null

            Write-Verbose "Starting mirroring session on primary server"
            $setPartnerPrimary = "ALTER DATABASE $Database SET PARTNER = 'TCP://$($MirrorServerInstance.Split('\')[0]):$($mirrorEndpoint.port)'"
            Send-SqlNonQuery $primaryConn $setPartnerPrimary | Out-Null

            if ($WitnessServerInstance) {
                Write-Verbose "Adding witness server to mirroring session"
                $setPartnerWitness = "ALTER DATABASE $Database SET WITNESS = 'TCP://$($WitnessServerInstance.Split('\')[0]):$($witnessEndpoint.port)'"
                Send-SqlNonQuery $primaryConn $setPartnerWitness | Out-Null
            }
        } catch {
            Write-Error $_
            return
        }
    }

    END {
        Close-SqlConnection $primaryConn
        Close-SqlConnection $mirrorConn
        Close-SqlConnection $witnessConn
    }
}

function Confirm-DatabaseMirroringReadiness {
    param (
            [ValidateNotNull()]
            [Object]
            $PrimaryDatabase,

            [ValidateNotNull()]
            [Object]
            $MirrorDatabase
          )

    $errors = @()

    if ($PrimaryDatabase.collation_name -ne $MirrorDatabase.collation_name) {
        $errors += "Collation mismatch:  Ensure the database collation attributes are identical on the primary and mirror servers."
    }

    if ($PrimaryDatabase.recovery_model -ne 1) {
        # 1 means FULL recovery model.
        $errors += "Primary database must be in the FULL recovery model, not $($PrimaryDatabase.recovery_model_desc), in order to configure mirroring."
    }

    if ($MirrorDatabase.recovery_model -ne 1) {
        $errors += "Mirror database must be in the FULL recovery model, not $($MirrorDatabase.recovery_model_desc), in order to configure mirroring."
    }

    if ($PrimaryDatabase.state_desc -ne "ONLINE") {
        $errors += "Primary database must be in the ONLINE state, not the $($PrimaryDatabase.state_desc) state."
    }

    if ($MirrorDatabase.state_desc -ne "RESTORING") {
        $errors += "Mirror database must be in the RESTORING state, not the $($MirrorDatabase.state_desc) state."
    }

    if ($errors) {
        foreach ($error in $errors) {
            Write-Error $error
        }
        return $false
    }
    return $true
}

function New-MirroringEndpoint {
    param (
            [ValidateNotNull()]
            [System.Data.SqlClient.SqlConnection]
            $SqlConnection,

            [ValidateRange(1025, 65536)]
            [int]
            $EnpointPort,

            [switch]
            $IsWitnessOnly
          )

    # "4" is a DATABASE_MIRRORING endpoint
    $cmd = "SELECT name, port, endpoint_id FROM sys.tcp_endpoints WHERE type = 4"
    $endpoint = @(Send-SqlQuery $SqlConnection $cmd)
    if ($endpoint) {
        Write-Warning "A database mirroring endpoint already exists for $($SqlConnection.Datasource). (Name: $($endpoint[0].name), Port: $($endpoint[0].port))  This endpoint will be used instead."
        return $endpoint
    }

    Write-Verbose "No pre-existing endpoints.  A new endpoint will be created."

    if ($IsWitnessOnly) {
        $role = "WITNESS"
    } else {
        $role = "ALL"
    }
    $createEndpointCmd = "CREATE ENDPOINT Mirroring STATE = STARTED AS TCP ( LISTENER_PORT = $EnpointPort ) FOR DATABASE_MIRRORING ( ROLE = $role )"
    Send-SqlNonQuery $SqlConnection $createEndpointCmd | Out-Null

    $endpoint = @(Send-SqlQuery $SqlConnection $cmd)
    if ($endpoint -eq $null) {
        throw "The database mirroring endpoint could not be created for $($SqlConnection.Datasource)."
    }

    return $endpoint
}

function Grant-PrivilegesOnEndpoint {
    param (
            [ValidateNotNull()]
            [System.Data.SqlClient.SqlConnection]
            $SqlConnection,

            [ValidateNotNullOrEmpty()]
            [string]
            $DatabaseUser,

            [ValidateNotNullOrEmpty()]
            [Object]
            $Endpoint
          )

    $getLogins = "SELECT name, principal_id FROM sys.server_principals WHERE name = '$DatabaseUser'"
    $login = Send-SqlQuery $SqlConnection $getLogins

    if ($login) {
        Write-Verbose "Server login for user $DatabaseUser on instance $($SqlConnection.Datasource) already exists."
    } else {
        Write-Verbose "Creating login for user $DatabaseUser on instance $($SqlConnection.Datasource)"
        $createLogin = "CREATE LOGIN [$DatabaseUser] FROM WINDOWS"
        Send-SqlNonQuery $SqlConnection $createLogin | Out-Null

        $login = Send-SqlQuery $SqlConnection $getLogins
        if ($login -eq $null) {
            throw "Could not create login for user $DatabaseUser on instance $($SqlConnection.Datasource)"
        }
    }

    # "105" is the ENDPOINT class.
    $checkPrivs = "SELECT COUNT(*) FROM sys.server_permissions WHERE grantee_principal_id = $($login.principal_id) AND class = 105 AND major_id = $($Endpoint.endpoint_id) AND type = 'CO'"
    $privCount = Send-SqlScalarQuery $SqlConnection $checkPrivs
    if ($privCount -gt 0) {
        Write-Verbose "CONNECT privileges already exist on the $($Endpoint.name) endpoint for user $DatabaseUser on instance $($SqlConnection.Datasource)"
    } else {
        $grantPrivs = "GRANT CONNECT ON ENDPOINT::$($Endpoint.name) TO [$DatabaseUser]"
        Send-SqlNonQuery $SqlConnection $grantPrivs | Out-Null
        $privCount = Send-SqlScalarQuery $SqlConnection $checkPrivs
        if ($privCount -eq 0) {
            throw "Could not grant CONNECT privileges to $DatabaseUser on instance $($SqlConnection.Datasource), endpoint $($Endpoint.name)"
        }
    }
}

