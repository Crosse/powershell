function Connect-RemoteServer {
    [CmdletBinding()]
    param (
            [Parameter(Mandatory=$true)]
            [ValidateNotNullOrEmpty()]
            [string]
            $ComputerName,

            [switch]
            $UseExistingSession,

            [Parameter(Mandatory=$false)]
            [System.Management.Automation.PSCredential]
            $Credential,

            [Parameter(Mandatory=$false)]
            [System.Management.Automation.Runspaces.AuthenticationMechanism]
            $Authentication,

            [switch]
            $UseSSL
          )

    if ($UseExistingSession) {
        $session = Get-PSSession $ComputerName -ErrorAction SilentlyContinue
    }
    if ($session -eq $null) {
        $cmd = 'New-PSSession -ComputerName $ComputerName'
        if ($Credential) {
            $cmd += ' -Credential $Credential'
        }
        if ($Authentication) {
            $cmd += ' -Authentication $Authentication'
        }
        if ($UseSSL) {
            $cmd += ' -UseSSL'
        }

        $session = Invoke-Expression -Command $cmd
        if ($session -eq $null) {
            Write-Error "Could not create session"
            return
        }
    } else {
        Write-Verbose "Using existing PSSession"
    }

    if (Test-Path $PROFILE) {
        $pInfo = [IO.FileInfo]$PROFILE
        $p = Get-Content $PROFILE
        $encoder = New-Object System.Text.ASCIIEncoding
        $sha = [System.Security.Cryptography.SHA256]::Create()
        $hash = [Convert]::ToBase64String($sha.ComputeHash($encoder.GetBytes($p)))

        $verifyProfileScript = {
            param ($localProfileName, $localProfileHash, $Verbose)
            if ($Verbose) {
                $VerbosePref = $VerbosePreference
                $VerbosePreference = $Verbose
            }
            Write-Verbose "Local profile hash:  $localProfileHash"

            $userProfilePath = (Get-Item Env:\USERPROFILE).Value
            $psProfilePath = Join-Path $userProfilePath "Documents\WindowsPowerShell"
            $PROFILE = Join-Path $psProfilePath $localProfileName
            $retval = $false
            if (Test-Path $PROFILE) {
                $p = Get-Content $PROFILE
                $encoder = New-Object System.Text.ASCIIEncoding
                $sha = [System.Security.Cryptography.SHA256]::Create()
                $hash = [Convert]::ToBase64String($sha.ComputeHash($encoder.GetBytes($p)))
                Write-Verbose "Remote profile hash: $hash"
                $retval = ($hash -eq $localProfileHash)
            } else {
                if (Test-Path $psProfilePath) {
                    Write-Verbose "Found $psProfilePath directory on $Env:ComputerName"
                } else {
                    Write-Verbose "Creating $psProfilePath directory on $Env:ComputerName"
                    try {
                        $null = New-Item -ItemType Directory -Path $psProfilePath
                    } catch {
                        Write-Warning "Could not create $psProfilePath on $Env:ComputerName"
                        $retval = $null
                    }
                }
            }
            if ($Verbose) {
                $VerbosePreference = $VerbosePref
            }
            Remove-Variable userProfilePath, psProfilePath
            return $retval
        }

        $copyProfileScript = {
            param ($localProfileName, $Verbose)
            if ($Verbose) {
                $VerbosePref = $VerbosePreference
                $VerbosePreference = $Verbose
            }

            $userProfilePath = (Get-Item Env:\USERPROFILE).Value
            $psProfilePath = Join-Path $userProfilePath "Documents\WindowsPowerShell"
            $PROFILE = Join-Path $psProfilePath $localProfileName

            Write-Verbose "Copying $localProfileName to $Env:ComputerName"
            try {
                $input | Out-File $PROFILE
                if (Test-Path $PROFILE) {
                    Write-Verbose "Successfully copied profile to $Env:ComputerName"
                }
            } catch {
                Write-Warning "Could not copy profile to $Env:ComputerName"
            }
            if ($Verbose) {
                $VerbosePreference = $VerbosePref
            }
            Remove-Variable userProfilePath, psProfilePath
        }

        $retval = Invoke-Command -Session $session `
                                 -ScriptBlock $verifyProfileScript `
                                 -ArgumentList $pInfo.Name, $hash, $VerbosePreference

        if ($retval -eq $false) {
            Get-Content $PROFILE |
                Invoke-Command `
                    -Session $session `
                    -ScriptBlock $copyProfileScript `
                    -ArgumentList $pInfo.Name, $VerbosePreference
        }

        Invoke-Command -Session $session { . $PROFILE }
        Enter-PSSession $session
    }
}

Set-Alias -Name go -Value Connect-RemoteServer -Scope Global
