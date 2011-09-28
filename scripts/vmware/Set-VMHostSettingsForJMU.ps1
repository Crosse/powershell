param (
        [Parameter(Mandatory=$false)]
        $VIServer,

        [Parameter(Mandatory=$true,
            ValueFromPipeline=$true)]
        [ValidateNotNullOrEmpty()]
        $VMHost,

        [Parameter(Mandatory=$true)]
        $VSphereCredential,

        [Parameter(Mandatory=$true)]
        $HostCredential
      )

function WriteWarning($message) {
    Write-Warning "$($server):`t$message"
}

function WriteHost($message) {
    Write-Host "$($server):`t$message"
}

function WriteError($message) {
    Write-Host "$($server):`t$message"
}

$vSphereServer = $null

if ($VIServer -ne $null) {
    if ($VIServer -eq $DefaultVIServer.Name) {
        Write-Verbose "Using existing connection to $DefaultVIServer"
        $vSphereServer = $DefaultVIServer
    } else {
        Write-Verbose "Connecting to $VIServer"
        $vSphereServer = Connect-VIServer $VIServer -Credential $VSphereCredential
    }
} else {
    $vSphereServer = Connect-VIServer -Menu
}

if ($vSphereServer -eq $null) {
    Write-Error "No VIServer specified."
    return
}

$server = Get-VMHost $VMHost
if ($server -eq $null) {
    Write-Error "Could not find host $VMHost"
    return
}

# Check for OpenManage
$patches = Get-VMHostPatch -Server $vSphereServer -VMHost $Server
$foundOM = $false
$foundDellImage = $false

WriteHost "Updating esxupdate.conf"
$svc = Get-VMHostService $server | ? { $_.Key -eq 'TSM-SSH' }
if ($svc.Running -eq $false) {
    $svc | Start-VMHostService
    Start-Sleep 5
}
& 'C:\Program Files (x86)\PuTTY\plink.exe' -l $HostCredential.GetNetworkCredential().UserName -pw $HostCredential.GetNetworkCredential().Password $server "sed -re 's|file = ?$|file = /var/tmp/esxupdate.debug|' -i /etc/vmware/esxupdate/esxupdate.conf"

foreach ($patch in $patches) {
    if ($patch.Id -match "OpenManage") {
        WriteHost "Found installed patch $($patch.Id) ($($patch.Description))"
        $foundOM = $true
    } elseif ($patch.Id -eq 'Dell') {
        WriteHost "Dell Customized ESXi image found"
        $foundDellImage = $true
    }
}

if ($foundDellImage -eq $false) {
    WriteWarning "host was not built using the Dell-customized ESXi image"
}

if ($foundOM -eq $false) {
    WriteWarning "Dell OpenManage not installed.  Please use VMware Update Manager to install."
#    $ds = Get-Datastore -Server $vSphereServer | ? { $_.Name -eq 'NAS:virtcluster-Mass-utility01' }
#    if ($ds -eq $null) {
#        WriteError "Could not find NAS:virtcluster-Mass-utility01 iSCSI LUN!"
#        WriteError "DellOM installation failed."
#    } else {
#        $dellOMPath = "/vmfs/volumes/" + $ds.Name + "/Dell_OpenManage_ESXi410_OM650/metadata.zip"
#        Install-VMHostPatch -Server $vSphereServer -VMHost $server -HostPath $dellOMPath
#
#        WriteHost "Enabling CIM OEM Providers"
#        Set-VMHostAdvancedConfiguration -Server $vSphereServer -VMHost $server -Name UserVars.CIMoemProviderEnabled -Value 1
#    }
} #else {
    WriteHost "Enabling CIM OEM Providers"
    Set-VMHostAdvancedConfiguration -Server $vSphereServer -VMHost $server -Name UserVars.CIMoemProviderEnabled -Value 1
#}

$ntpServers = @(Get-VMHostNtpServer -Server $vSphereServer -VMHost $server)
foreach ($ntpServer in $ntpServers) {
    WriteHost "Removing existing NTP Server $ntpServer"
    Remove-VmHostNtpServer -Server $vSphereServer -VMHost $server -NtpServer $ntpServer -Confirm:$false
}
$gateway = (Get-VMHostNetwork -Server $vSphereServer -VMHost $server).VMKernelGateway
WriteHost "Adding NTP Server $gateway"
$null = Add-VMHostNtpServer -Server $vSphereServer -VMHost $server -NtpServer $gateway

WriteHost "Enabling syslog to it-lms"
$null = Set-VMHostSysLogServer -Server $vSphereServer -VMHost $server -SysLogServer it-lms.jmu.edu -SysLogServerPort 514

WriteHost "Enabling SNMP"
$null = Connect-VIServer -Server $server -Credential $HostCredential
$hostSnmp = Get-VMHostSnmp
foreach ($trapTarget in $hostSnmp.TrapTargets) {
    $hostSnmp = Set-VMHostSnmp -HostSnmp $hostSnmp -TrapTargetToRemove $trapTarget
}
$hostSnmp = Set-VMHostSnmp -HostSnmp $hostSnmp -ReadOnlyCommunity @("zenosscmnty", "hostwatch") 
$hostSnmp = Set-VMHostSnmp -HostSnmp $hostSnmp -AddTarget -TargetCommunity "zenosscmnty" -TargetHost "it-zenoss1.jmu.edu" -TargetPort 161
$hostSnmp = Set-VMHostSnmp -HostSnmp $hostSnmp -AddTarget -TargetCommunity "hostwatch" -TargetHost "ita.ad.jmu.edu" -TargetPort 161
$hostSnmp = Set-VMHostSnmp -HostSnmp $hostSnmp -Enabled $true

