<#
    .SYNOPSIS
    Gets Office 365 unified audit logs in a restartable, self-contained fashion.

    .DESCRIPTION
    As per https://msdn.microsoft.com/en-us/office-365/office-365-management-activity-api-reference, it can take up to 12 hours for log events to become available due to the distributed nature of the Office 365 servers and datacenters. This script simply scans a path containing previously-downloaded management activity log files, finds the most recently-written file, and calls Get-ManagementActivityLogs (in Crosse.PowerShell.Office365; see https://github.com/crosse/powershell/) with the -Start parameter set to twelve hours prior to the msot recent log file it found. (This doesn't guarantee that all events are eventually returned, but it's probably mostly correct...)

    .EXAMPLE
    C:\PS> $cred = Get-Credential; .\GetO365Logs.ps1 -Credential $cred -Path C:\ManagementActivityLogs
    Requesting logs from 09/15/2016 03:00:00 until 09/15/2016 17:00:33

    This example shows saving Office 365 credentials in the $cred variable, then passing it to the GetO365Logs.ps1 script. The most recent log file prior to running this script appears to have been created somewhere between 3:00PM and 4:00PM on Sept. 15, 2016.
#>

#Requires -Version 5.0
#Requires -Modules @{ModuleName="Crosse.PowerShell.Office365"; ModuleVersion="1.0"; GUID='dd3c22d1-9016-4e2d-bfda-9b98588d0540'}
[CmdletBinding()]
param (
    [Parameter(Mandatory=$false)]
    [System.Management.Automation.PSCredential]
    [System.Management.Automation.Credential()]
    # A System.Managemement.Automation.PSCredential credential for an Office 365 user able to use the Search-UnifiedAuditLog cmdlet. If this parameter is not specified, then the script will attempt to find an open PSSession to outlook.office365.com and use that instead.
    $Credential = [System.Management.Automation.PSCredential]::Empty,

    [Parameter(Mandatory=$true)]
    [string]
    # The base directory where log files should be written.
    $Path
)

# Get the creation time of the last log file written...
$start = (Get-ChildItem -Recurse $Path -Filter *.csv | Sort CreationTime -Desc)[0].CreationTime
# ...and then set the start time to twelve hours before that, on the hour.
$start = $start.AddSeconds(-$start.Second).AddMinutes(-$start.Minute).AddHours(-12)
$start = $start.AddTicks(-($start.Ticks % 10000000))

# PowerShell 5.0 sets $InformationPreference to "SilentlyContinue" by default.
$InformationPreference = "Continue"

Write-Information -MessageData "Requesting logs from $start until $(Get-Date)"

# If there is an existing session to Office 365, use that.
$Session = @(Get-PSSession | ? { $_.ComputerName -eq "outlook.office365.com" -and $_.State -eq 'Opened' })[0]
if ($Session -eq $null) {
    # Otherwise, construct a session using $Credential, if available.
    if ($Credential -eq [System.Management.Automation.PSCredential]::Empty) {
        Write-Error "No open Office 365 session and no credentials given!"
        exit
    }

    $Session = New-Office365Session -Credential $Credential -ConnectToAzureAD:$false -ImportSession:$false
    if ($session -eq $null) {
        Write-Error "Session creation failed!"
        exit
    }
}
Get-ManagementActivityLogs -Start $start -Path $Path -Session $Session