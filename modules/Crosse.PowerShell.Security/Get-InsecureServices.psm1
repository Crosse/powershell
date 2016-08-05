function Get-InsecureServices {
    [CmdletBinding()]
    param (
            [Parameter(Mandatory=$false)]
            [ValidateNotNullOrEmpty()]
            [string]
            $ComputerName = "localhost"
          )

    $everyoneSid = ([System.Security.Principal.NTAccount]"Everyone").Translate([System.Security.Principal.SecurityIdentifier])
    $rights = @{
        SERVICE_CHANGE_CONFIG = 0x00000002
        WRITE_DAC = 0x00040000
        WRITE_OWNER = 0x00080000
    }
    $rightsToCheck = 0
    $rights.Values | Foreach-Object { $rightsToCheck = $rightsToCheck -bor $_ }

    if ([String]::IsNullOrEmpty($ComputerName)) {
        $services = Get-Service
    } else {
        $services = Get-Service -ComputerName $ComputerName
    }

    Write-Verbose "Found $($services.Count) services"

    for ($i = 0; $i -lt $services.Count; $i++) {
        $service = $services[$i]
        [Int]$pctComplete = [Math]::Floor($i / $services.Count * 100)
        Write-Progress -Activity "Evaluating services" -Status "$pctComplete% Complete" -CurrentOperation $service.Name -PercentComplete $pctComplete
        $sddl = sc.exe "\\$ComputerName" sdshow $service.Name | Where-Object { $_ }

        try {
            $sd = New-Object System.Security.AccessControl.CommonSecurityDescriptor($false, $false, $sddl)
        } catch {
            Write-Warning ("Error creating security descriptor for {0}: {1}" -f $service.Name, $_.Exception.Message)
            continue
        }

        if ($sd.DiscretionaryAcl | Where-Object {
                $_.AceQualifier -eq [System.Security.AccessControl.AceQualifier]::AccessAllowed `
                -and $_.SecurityIdentifier -eq $everyoneSid `
                -and $_.AccessMask -band $RightsToCheck }) {
            Write-Warning "Found insecure service: $($service.Name)"
            $service
        }
    }
    Write-Progress -Activity "Evaluating services" -Status "Progress:" -Completed
}
