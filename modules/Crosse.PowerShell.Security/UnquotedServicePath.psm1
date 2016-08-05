function Test-ServicePath {
    [CmdletBinding()]
    param (
            [Parameter(Mandatory=$false,
                ValueFromPipeline=$true)]
            [ValidateNotNullOrEmpty()]
            [string]
            $ComputerName = (Get-Item Env:\COMPUTERNAME).Value,

            [Parameter(Mandatory=$false)]
            [ValidateNotNullOrEmpty()]
            [string[]]
            $Service = $null
          )

    BEGIN {
        Set-StrictMode -Version Latest
    }
    PROCESS {
        if ($Service) {
            $services = Get-WmiObject -ComputerName $ComputerName `
                        -Class Win32_Service -Filter "Name = '$Service'" | Sort-Object Name
        } else {
            $services = Get-WmiObject -ComputerName $ComputerName `
                        -Class Win32_Service | Sort-Object Name
        }

        $paths = @{}
        for ($j = 0; $j -lt $services.Count; $j++) {
            $svc = $services[$j]
            [Int]$pct = ($j/$services.Count) * 100
            Write-Progress -Activity "Finding insecure service paths on $ComputerName" `
                        -Status $svc.Name -PercentComplete $pct
            $svcName = $svc.Name
            $path = $svc.PathName.Trim()

            if ($path.StartsWith('"')) {
                # This service path is already (hopefully) properly quoted.
                # Extract the command (which should be the first part of the
                # string in quotes) and the arguments, the ship the command
                # off to check whether it's actually a valid command.
                $commandEndIndex = $path.IndexOf('"', 1)
                $command = $path.Substring(1, $commandEndIndex - 1)
                $arguments = $path.Substring($commandEndIndex + 1).Trim()

                Write-Verbose "[$svcName] - Quoted path found. Testing command $command"
                # First check the cache.
                if ($paths.Keys -contains $command) {
                    $file = $paths[$command]
                } else {
                    $file = FindCommand -ComputerName $ComputerName -Command $command
                    $paths[$command] = $file
                }

                if ($file) {
                } else {
                    Write-Error "[$svcName] - Quoted path found, but the command was not: $path"
                }
                # Whatever the outcome, we're done with this service.
                continue
            }

            if ($path -notmatch '\s') {
                # Not handling this case yet.
                Write-Verbose "[$svcName] - Command does not contain any spaces."
                continue
            }

            # Getting here means that the path is unquoted.
            $pieces = @($path.Split(" `t", [StringSplitOptions]::RemoveEmptyEntries) | Foreach-Object { $_.Trim() })
            $found = $false

            for ($i = 0; $i -lt $pieces.Count; $i++) {
                $command = $pieces[0..$i] -join " "

                # First check the cache.
                if ($paths.Keys -contains $command) {
                    $file = $paths[$command]
                } else {
                    $file = FindCommand -ComputerName $ComputerName -Command $command
                    $paths[$command] = $file

                    if ($file -and $command -ine $file.Name) {
                        Write-Warning "[$svcName] - Service's command and actual path do not match: `"$command`" vs. `"$($file.Name)`""
                        $command = $file.Name
                    }
                }

                if ($file) {
                    $found = $true
                    break
                }
            }

            if ($found) {
                if ($command -match '\s') {
                    Write-Verbose "[$svcName] - Path needs to be quoted ($command)"
                    New-Object PSObject -Property @{
                        ComputerName= $ComputerName
                        Name        = $svcName
                        PathName    = $path
                        Command     = $command
                        Arguments   = $path.Replace($command, "").Trim()
                        Service     = (Get-Service -ComputerName $ComputerName -Name $svcName)
                    }
                }
            } else {
                Write-Error "[$svcName] - No valid command found in service path ($path)"
            }
        }
    }
}

function Repair-ServicePath {
    [CmdletBinding(SupportsShouldProcess=$true,
            ConfirmImpact="High")]
    param (
            [Parameter(Mandatory=$true,
                ValueFromPipelineByPropertyName)]
            [System.ServiceProcess.ServiceController]
            $Service,

            [Parameter(Mandatory=$false,
                ValueFromPipelineByPropertyName)]
            [String]
            $ComputerName = (Get-Item Env:\COMPUTERNAME).Value,

            [Parameter(Mandatory=$true,
                ValueFromPipelineByPropertyName)]
            [String]
            $Command,

            [Parameter(Mandatory=$true,
                ValueFromPipelineByPropertyName)]
            [AllowEmptyString()]
            [String]
            $Arguments
          )

    BEGIN {
        Set-StrictMode -Version Latest
    }
    PROCESS {
        $svcName = $Service.Name

        $svcWmiInstance = Get-WmiObject -ComputerName $ComputerName -Class Win32_Service -Filter "Name = '$($Service.Name)'"
        if ($svcWmiInstance -eq $null) {
            Write-Error "Cound not find the requested service!"
            return
        }

        $binPath = "`"{0}`"" -f $Command.Trim()
        if (![String]::IsNullOrEmpty($Arguments)) {
            $binPath += " " + $Arguments.Trim()
        }
        if ($PSCmdlet.ShouldProcess("Repairing unquoted path for the `"$svcName`" service on $ComputerName",
                    "This operation will properly quote the path for the `"$svcName`" service.
    Old Path: $($svcWmiInstance.PathName)
    New Path: $binPath", "Repair Unquoted Path for `"$svcName`" service")) {
            Write-Verbose "Repairing service path name for `"$svcName`" on $ComputerName"
            $result = $svcWmiInstance.Change($null,$binPath,$null,$null,$null,$null,$null,$null,$null,$null,$null)
            if ($result.ReturnValue -ne 0) {
                Write-Error "An error occurred trying to modify the service's path!"
            } else {
                $svcWmiInstance = Get-WmiObject -ComputerName $ComputerName -Class Win32_Service -Filter "Name = '$($Service.Name)'"
                New-Object PSObject -Property @{
                    ComputerName= $ComputerName
                    Name        = $svcName
                    PathName    = $svcWmiInstance.PathName
                    Service     = (Get-Service -ComputerName $ComputerName -Name $svcName)
                }
            }
        }
    }
}

function FindCommand {
    [CmdletBinding()]
    param (
            [Parameter(Mandatory=$true)]
            [ValidateNotNullOrEmpty()]
            [String]
            $Command,

            [Parameter(Mandatory=$false)]
            [ValidateNotNullOrEmpty()]
            [String]
            $ComputerName = (Get-Item Env:\COMPUTERNAME).Value
          )

    BEGIN {
        Set-StrictMode -Version Latest
    }

    PROCESS {
        Write-Verbose "Command: $Command"

        # WMI requires escaping of the backslashes
        $escapedPath = $Command.Replace('\', '\\')

        # First, just see if the path as-presented is the path to a real file.
        $file = Get-WmiObject -ComputerName $ComputerName `
                -Class CIM_DataFile -Filter "Name = '$escapedPath'"

        if ($file -ne $null) {
            return $file
        }

        $found = $false
        $pathext = (Get-Item Env:\PATHEXT).Value.Split(';')
        foreach ($ext in $pathext) {
            # We couldn't find a valid file using the path as-is, so start
            # tacking on (in order of precedence) valid command extensions to see
            # if we get a hit.
            $escapedPath = $Command.Replace('\', '\\') + $ext
            $file = Get-WmiObject -ComputerName $ComputerName `
                    -Class CIM_DataFile -Filter "Name = '$escapedPath'"

            if ($file) {
                Write-Verbose "Found file by appending ${ext}: $($file.Name)"
                return $file
            }
        }
        Write-Warning "Command not found: $Command"
    }
}

Export-ModuleMember -Function Test-ServicePath, Repair-ServicePath
