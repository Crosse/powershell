function Set-FirewallRule {
    [CmdletBinding()]
    param (
            [Parameter(Mandatory=$true)]
            [ValidateNotNullOrEmpty()]
            [string]
            $Name,

            [ValidateSet("In", "Out")]
            [string]
            $Direction="In",

            [ValidateSet("Allow", "Block", "Bypass")]
            [string]
            $Action="Allow",

            [ValidateNotNullOrEmpty()]
            [string]
            $Description,

            [switch]
            $Enable=$true,

            [ValidateSet("Public", "Private", "Domain", "Any")]
            [string[]]
            $Profile,

            [ValidateNotNullOrEmpty()]
            [string]
            $Protocol="TCP",

            [ValidateNotNullOrEmpty()]
            [string[]]
            $LocalPorts,

            [ValidateNotNullOrEmpty()]
            [string]
            $LocalIp="any",

            [Parameter(Mandatory=$true)]
            [ValidateNotNullOrEmpty()]
            [string[]]
            $RemoteIps,

            [switch]
            $Force
          )

    if ($Enable) {
        $en = "yes"
    } else {
        $en = "no"
    }

    $pro = [String]::Join(",", $Profile)
    $localport = [String]::Join(",", $LocalPorts)
    $remoteip = [String]::Join(",", $RemoteIps)

    netsh.exe advfirewall firewall show rule name="$Name" | Out-Null
    if ($? -eq $false) {
        Write-Verbose "Rule `"$Name`" not found.  Rule will be created."
        $result = netsh.exe advfirewall firewall add rule `
            name="$Name" `
            dir=$Direction `
            action=$Action `
            description="$Description" `
            enable=$en `
            profile="$pro" `
            protocol="`"$Protocol`"" `
            localport="$localport" `
            localip="`"$LocalIp`"" `
            remoteip="$remoteip"
        if ($?) {
            Write-Host "Rule `"$Name`" created."
        } else {
            $err = @()
            foreach ($line in $result) {
                if ([String]::IsNullOrEmpty($line)) { continue }
                if ($line -match "Usage") { break }
                $err += $line
            }
            Write-Error "Error creating rule `"$Name`":  $err"
        }
    } else {
        if ($Force) {
            Write-Verbose "Updating rule `"$Name`"."
            $result = netsh.exe advfirewall firewall set rule `
                name="$Name" `
                dir=$Direction `
                new `
                action=$Action `
                description="$Description" `
                enable=$en `
                profile="$pro" `
                protocol="`"$Protocol`"" `
                localport="$localport" `
                localip="`"$LocalIp`"" `
                remoteip="$remoteip"

            if ($?) {
                Write-Host "Rule `"$Name`" updated."
            } else {
                $err = @()
                foreach ($line in $result) {
                    if ([String]::IsNullOrEmpty($line)) { continue }
                    if ($line -match "Usage") { break }
                    $err += $line
                }
                Write-Error "Error updating rule `"$Name`":  $err"
            }
        } else {
            Write-Error "Rule `"$Name`" exists, and -Force was not specified.  Not overwriting."
        }
    }
}
