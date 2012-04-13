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

            [Parameter(Mandatory=$false)]
            [switch]
            $ShowProgress,

            [Parameter(Mandatory=$false,
                ParameterSetName="TransactionLogBackup")]
            [switch]
            $Truncate
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
            $withOption += "NOINIT"
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

    if ($BoundParameters.ContainsKey('TRUNCATE')) {
        if ($Truncate -eq $false) {
            $option = "NO_TRUNCATE"
            Write-Verbose "Requested: $option"
            $withOptions += $option
        }
    }

    return $withOptions
}

Export-ModuleMember Backup-Database
