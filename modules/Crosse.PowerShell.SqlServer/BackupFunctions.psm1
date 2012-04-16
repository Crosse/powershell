function Backup-Database {
    [CmdletBinding()]
    param (
            [Parameter(Mandatory=$true)]
            [ValidateNotNullOrEmpty()]
            [string]
            $Server,

            [Parameter(Mandatory=$true)]
            [ValidateNotNullOrEmpty()]
            [string]
            $Database,

            [Parameter(Mandatory=$true)]
            [System.IO.FileInfo]
            $BackupPath,

            [Parameter(Mandatory=$false,
                ParameterSetName="FullBackup")]
            [switch]
            $FullBackup,

            [Parameter(Mandatory=$false,
                ParameterSetName="DifferentialBackup")]
            [switch]
            $DifferentialBackup,

            [Parameter(Mandatory=$false,
                ParameterSetName="TransactionLogBackup")]
            [switch]
            $TransactionLogBackup,

            [Parameter(Mandatory=$false)]
            [switch]
            $CopyOnly,

            [Parameter(Mandatory=$false)]
            [switch]
            $Compression,

            [Parameter(Mandatory=$false)]
            [string]
            $Description,

            [Parameter(Mandatory=$false)]
            [string]
            $Name,

            [Parameter(Mandatory=$false)]
            [DateTime]
            $ExpireDate,

            [Parameter(Mandatory=$false)]
            [int]
            $RetainDays,

            [Parameter(Mandatory=$false)]
            [switch]
            $Init,

            [Parameter(Mandatory=$false)]
            [switch]
            $Skip,

            [Parameter(Mandatory=$false)]
            [switch]
            $Format,

            [Parameter(Mandatory=$false)]
            [string]
            $MediaDescription,

            [Parameter(Mandatory=$false)]
            [string]
            $MediaName,

            [Parameter(Mandatory=$false)]
            [switch]
            $Checksum,

            [Parameter(Mandatory=$false,
                ParameterSetName="TransactionLogBackup")]
            [switch]
            $NoTruncate
          )

    if ($FullBackup -or $DifferentialBackup) {
        $backupType = "DATABASE"
    } elseif ($TransactionLogBackup) {
        $backupType = "LOG"
    } else {
        # should never happen
        Write-Error "Backup type not recognized"
        return
    }

    $cmd = "BACKUP {0} {1} TO DISK = '{2}'" -f $backupType, $Database, $BackupPath

    $withOptions = @(BuildWithOptions $PSCmdlet.MyInvocation.BoundParameters)
    if ($withOptions.Count -gt 0) {
        $with = " WITH {0}" -f ($withOptions -join ", ")
        $cmd += $with
    }

    Write-Verbose $cmd
    $conn = Open-SqlConnection -Server $Server -Database $Database -Async
    $spid = (Send-SqlQuery -SqlConnection $conn -Command "SELECT @@SPID as spid").spid
    Write-Verbose "SPID: $spid"
    $sqlCmd = $conn.CreateCommand()
    $sqlCmd.CommandText = $cmd

    $checkConn = Open-SqlConnection -Server $Server
    $perms = Send-SqlQuery -SqlConnection $checkConn -Command "SELECT * FROM fn_my_permissions(NULL, 'SERVER') WHERE permission_name = 'VIEW SERVER STATE'"
    if ($perms -eq $null) {
        Write-Warning "Cannot get backup progress.  You have not been granted the VIEW SERVER STATE permission."
    }

    $start = Get-Date
    $result = $sqlCmd.BeginExecuteNonQuery()
    while (! $result.IsCompleted) {
        $check = Send-SqlQuery -SqlConnection $checkConn -Command "SELECT session_id,percent_complete,command,estimated_completion_time FROM sys.dm_exec_requests WHERE session_id = $spid AND command = 'BACKUP DATABASE'"
        if ($check -eq $null) {
            break
        }

        if ($perms -eq $null) {
            $elapsed = ((Get-Date) - $start).ToString("hh\:mm\:ss")
            Write-Progress -Activity "Backing up $Database" -Status "Elapsed time:  $elapsed"
        } else {
            $percent = [Math]::Round($check.percent_complete, 0, "AwayFromZero")
            Write-Progress -Activity "Backing up $Database" `
                           -Status "${percent}% complete" `
                           -PercentComplete $check.percent_complete `
                           -SecondsRemaining ($check.estimated_completion_time / 1000)
        }
        Start-Sleep -Milliseconds 250
    }
    Write-Progress -Activity "Backing up $Database" -Status "Completed" -Completed
    $sqlCmd.EndExecuteNonQuery($result) | Out-Null
    $sqlCmd.Dispose()
    Close-SqlConnection $conn
    Close-SqlConnection $checkConn
}

function BuildWithOptions {
    param (
            [System.Collections.Hashtable]
            [ValidateNotNull()]
            $BoundParameters
          )

    $withOptions = @()

    if ($BoundParameters['DifferentialBackup']) {
        $option = "DIFFERENTIAL"
        Write-Verbose "Requested: $option"
        $withOptions += $option
    }

    if ($BoundParameters['CopyOnly']) {
        $option = "COPY_ONLY"
        Write-Verbose "Requested: $option"
        $withOptions += $option
    }

    if ($BoundParameters.ContainsKey('Compression')) {
        if ($Compression) {
            $option = "COMPRESSION"
            Write-Verbose "Requested: $option"
            $withOptions += $option
        } else {
            $option = "NO_COMPRESSION"
            Write-Verbose "Requested: $option"
            $withOptions += $option
        }
    }

    if (! [String]::IsNullOrEmpty($Description)) {
        $option = "DESCRIPTION = '$Description'"
        Write-Verbose "Requested: $option"
        $withOptions += $option
    }

    if (! [String]::IsNullOrEmpty($Name)) {
        $option = "NAME = '$Name'"
        Write-Verbose "Requested: $option"
        $withOptions += $option
    }

    if ($ExpireDate) {
        $option = "EXPIREDATE = '$ExpireDate'"
        Write-Verbose "Requested: $option"
        $withOptions += $option
    }

    if ($RetainDays) {
        $option = "RETAINDAYS = $RetainDays"
        Write-Verbose "Requested: $option"
        $withOptions += $option
    }

    if ($BoundParameters.ContainsKey('Init')) {
        if ($Init) {
            $option = "INIT"
            Write-Verbose "Requested: $option"
            $withOptions += $option
        } else {
            $option = "NOINIT"
            Write-Verbose "Requested: $option"
            $withOptions += $option
        }
    }

    if ($BoundParameters.ContainsKey('Skip')) {
        if ($Skip) {
            $option = "SKIP"
            Write-Verbose "Requested: $option"
            $withOptions += $option
        } else {
            $option = "NOSKIP"
            Write-Verbose "Requested: $option"
            $withOptions += $option
        }
    }

    if ($BoundParameters.ContainsKey('Format')) {
        if ($Format) {
            $option = "FORMAT"
            Write-Verbose "Requested: $option"
            $withOptions += $option
        } else {
            $option = "NOFORMAT"
            Write-Verbose "Requested: $option"
            $withOptions += $option
        }
    }

    if (! [String]::IsNullOrEmpty($MediaDescription)) {
        $option = "MEDIADESCRIPTION = '$MediaDescription'"
        Write-Verbose "Requested: $option"
        $withOptions += $option
    }

    if (! [String]::IsNullOrEmpty($MediaName)) {
        $option = "MEDIANAME = '$MediaName'"
        Write-Verbose "Requested: $option"
        $withOptions += $option
    }

    if ($BoundParameters.ContainsKey('Checksum')) {
        if ($Checksum) {
            $option = "CHECKSUM"
            Write-Verbose "Requested: $option"
            $withOptions += $option
        } else {
            $option = "NO_CHECKSUM"
            Write-Verbose "Requested: $option"
            $withOptions += $option
        }
    }

    if ($BoundParameters.ContainsKey('NoTruncate')) {
        if ($NoTruncate) {
            $option = "NO_TRUNCATE"
            Write-Verbose "Requested: $option"
            $withOptions += $option
        }
    }

    return $withOptions
}

Export-ModuleMember Backup-Database
