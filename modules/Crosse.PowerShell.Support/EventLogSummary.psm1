################################################################################
#
# Copyright (c) 2011 Seth Wright <seth@crosse.org>
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

################################################################################
<#
    .SYNOPSIS
    Gets events from event logs on local and remote computers and summarizes the results.

    .DESCRIPTION
    The Get-EventLogSummary cmdlet gets events from event logs such as the System
    and Application logs and summarizes them by deduplicating events and only
    returning the count of how many times each event occurred, along with other
    summary data.

    .INPUTS
    System.String.  You can pipe a ComputerName to Get-EventLogSummary.

    .OUTPUTS
    An array of PSObjects containing summary data for each event.

    .EXAMPLE
    PS C:\> Get-EventLogSummary -LogName Application -Start (Get-Date).AddDays(-1)


    Dedupid          : SceCli_1704
    Id               : 1704
    Count            : 1
    FirstTime        : 2/22/2012 7:01:20 AM
    Level            : 4
    TaskDisplayName  :
    ProviderName     : SceCli
    LogName          : Application
    MachineName      : localdesktop.contoso.com
    LevelDisplayName : Information
    SampleMessage    : Security policy in the Group policy objects has been applied successfully.
    LastTime         : 2/22/2012 7:01:20 AM

#>
################################################################################
function Get-EventLogSummary {
    param (
            [Parameter(Mandatory=$false,
                ValueFromPipeline=$true)]
            [string]
            # Gets events from the event logs on the specified computer.
            $ComputerName,

            [Parameter(Mandatory=$true)]
            [string[]]
            # Gets events from the specified event logs. Enter the event log
            # names in a comma-separated list. Wildcards are permitted
            $LogName,

            [Parameter(Mandatory=$false)]
            [DateTime]
            # Gets only the events that occur after the specified date and
            # time. Enter a DateTime object, such as the one returned by the
            # Get-Date cmdlet.
            $Start,

            [Parameter(Mandatory=$false)]
            [DateTime]
            # Gets only the events that occur before the specified date and
            # time. Enter a DateTime object, such as the one returned by the
            # Get-Date cmdlet.
            $End,

            [Parameter(Mandatory=$false)]
            [Int32[]]
            # Gets only the events that correspond to the specified log level(s).
            $Level,

            [switch]
            # Output the events as a single object down the pipeline instead of individually.
            $AsSingleObject
        )

    BEGIN {
        $eventParams = @{ LogName=$LogName }
        if ($Start -ne $null) {
            $eventParams.Add('StartTime', $Start)
        }
        if ($End -ne $Null) {
            $eventParams.Add('EndTime', $End)
        }

        if ($Level -ne $null) {
            $eventParams.Add('Level', $Level)
        }
    }
    PROCESS {
        if ([String]::IsNullOrEmpty($ComputerName)) {
            $ComputerName = "localhost"
        }
        Write-Verbose "Getting specified events for $ComputerName"

        $retval = @{}

        foreach ($event in (Get-WinEvent -ComputerName $ComputerName -FilterHashTable $eventParams -Verbose:$false)) {
            $dedupid = [String]::Format("{0}_{1}", $event.ProviderName.Replace(" ", "_") , $event.Id)

            if ($retval.Contains($dedupid)) {
                $retval[$dedupid].Count += 1
                if ($retval[$dedupid].FirstTime -gt $event.TimeCreated) {
                    $retval[$dedupid].FirstTime = $event.TimeCreated
                }
                if ($retval[$dedupid].LastTime -lt $event.TimeCreated) {
                    $retval[$dedupid].LastTime = $event.TimeCreated
                }
            } else {
                # This stuff doesn't work in PowerShell 2.0.
#                $defaultProperties = @('Count','ProviderName','Id','LevelDisplayName','SampleMessage')
#                $defaultDisplayPropertySet =
#                    New-Object System.Management.Automation.PSPropertySet(
#                            'DefaultDisplayPropertySet',
#                            [string[]]$defaultProperties)
#
#                $PSStandardMembers =
#                    [System.Management.Automation.PSMemberInfo[]]@($defaultDisplayPropertySet)
#
                $obj = New-Object PSObject -Property @{
                    Count               = 1
                    ProviderName        = $event.ProviderName
                    Id                  = $event.Id
                    Level               = $event.Level
                    LevelDisplayName    = $event.LevelDisplayName
                    LogName             = $event.LogName
                    MachineName         = $event.MachineName
                    TaskDisplayName     = $event.TaskDisplayName
                    FirstTime           = $event.TimeCreated
                    LastTime            = $event.TimeCreated
                    SampleMessage       = $event.Message
                    Dedupid             = $dedupid
                }

#                Add-Member -InputObject $obj -MemberType MemberSet `
#                                             -Name PSStandardMembers `
#                                             -Value $PSStandardMembers
                $retval[$dedupid] = $obj
            }
        }

        if ($AsSingleObject) {
            return ,@($retval.Values)
        } else {
            return @($retval.Values)
        }
    }
}

################################################################################
<#
    .SYNOPSIS
    Formats data generated by the Get-EventLogSummary cmdlet and sends the
    results via email.

    .DESCRIPTION
    Formats data generated by the Get-EventLogSummary cmdlet and sends the
    results via email.

    .INPUTS
    PSObject.  You can pipe the output of Get-EventLogSummary to
    Send-EventLogSummaryMailMessage.

    .OUTPUTS
    None.  Send-EventLogSummaryMailMessage sends an email to one or more
    recipients with the formatted data.
#>
################################################################################
function Send-EventLogSummaryMailMessage {
    [CmdletBinding(SupportsShouldProcess=$true,
            ConfirmImpact="High")]
    param (
            [Parameter(Mandatory=$true,
                ValueFromPipeline=$true)]
            [ValidateNotNullOrEmpty()]
            [PSObject[]]
            $Events,

            [Parameter(Mandatory=$true)]
            [ValidateNotNullOrEmpty()]
            [string]
            $SmtpServer,

            [Parameter(Mandatory=$true)]
            [ValidateNotNullOrEmpty()]
            [System.Net.Mail.MailAddress]
            $From,

            [Parameter(Mandatory=$true)]
            [ValidateNotNullOrEmpty()]
            [System.Net.Mail.MailAddress[]]
            $To,

            [Parameter(Mandatory=$false)]
            [ValidateNotNullOrEmpty()]
            [string]
            $Subject = "Event Log Summary Report for $(Get-Date -Format d)"
          )

    Add-Type -AssemblyName System.Web

    # This is ugly.
    $fontStyle      = "font-family:sans-serif;color:black;"
    $borderStyle    = "border:1px solid green;border-collapse:collapse;"
    $bodyStyle      = "$fontStyle;background-color:white;"
    $tableStyle     = "$bodyStyle;$borderStyle;margin-left:auto;margin-right:auto;width:90%;"
    $captionStyle   = "$fontStyle;border-style:none;border-collapse:collapse;text-align:center;font-weight:bold;font-size:1.2em;padding-left:6px;padding-right:6px;"
    $baseTrStyle    = "$fontStyle;$borderStyle;padding:6px;"
    $trStyle        = "$baseTrStyle;background-color:white;"
    $trAltStyle     = "$baseTrStyle;background-color:#EAF2D3;"
    $thStyle        = "$trStyle;$white-space:nowrap;text-align:center;"

    $trCriticalStyle    = "$baseTrStyle;background:rgba(255, 0, 0, 0.4);"
    $trCriticalAltStyle = "$baseTrStyle;background:rgba(255, 0, 0, 0.2);"

    $trErrorStyle       = "$baseTrStyle;background:rgba(255, 71, 25, 0.4);"
    $trErrorAltStyle    = "$baseTrStyle;background:rgba(255, 71, 25, 0.2);"

    $trWarningStyle     = "$baseTrStyle;background:rgba(255, 255, 143, 0.4);"
    $trWarningAltStyle  = "$baseTrStyle;background:rgba(255, 255, 143, 0.2);"

    $trInfoStyle        = "$baseTrStyle;background:rgba(0, 0, 255, 0.4);"
    $trInfoAltStyle     = "$baseTrStyle;background:rgba(0, 0, 255, 0.2);"


    $Body = @"
<!DOCTYPE html PUBLIC "-//W3C/DTD XHTML 1.0 Transitional//EN"
"http://www.w3.org/xhtml1/DTD/xhtml1-transitional.dtd">
<html>
<head>
    <title>$Subject</title>
</head>
<body style="$bodyStyle">
    <table style="$tableStyle">
        <caption style="$captionStyle">Event Log Summary Report</caption>
        <tr style="$trStyle">
            <th style="$thStyle">Server</th>
            <th style="$thStyle">EventId</th>
            <th style="$thStyle">Count</th>
            <th style="$thStyle">Log Level</th>
            <th style="$thStyle">Log name</th>
            <th style="$thStyle">Sample Message</th>
        </tr>
"@

    for ($i = 0; $i -lt $Events.Count; $i++) {
        if ($i % 2) {
            switch ($Events[$i].Level) {
                1 { $style = $trCriticalStyle; break }
                2 { $style = $trErrorStyle; break }
                3 { $style = $trWarningStyle; break }
                4 { $style = $trInfoStyle; break }
                default { $style = $trStyle; break }
            }
        } else {
            switch ($Events[$i].Level) {
                1 { $style = $trCriticalAltStyle; break }
                2 { $style = $trErrorAltStyle; break }
                3 { $style = $trWarningAltStyle; break }
                4 { $style = $trInfoAltStyle; break }
                default { $style = $trAltStyle; break }
            }
        }

        $message = $null
        $message = $Events[$i].SampleMessage
        if ($message.Length -gt 1024) {
            $message = $message.Substring(0, 1024)
            $message += " [...]"
        }
        $message = [System.Web.HttpUtility]::HtmlEncode($message)

        $Body += "
        <tr style=`"$style`">
            <td style=`"$style;white-space:nowrap;`">{0}</td>
            <td style=`"$style`">{1} {2}</td>
            <td style=`"$style;white-space:nowrap;text-align:right;`">{3}</td>
            <td style=`"$style;white-space:nowrap;`">{4}</td>
            <td style=`"$style;white-space:nowrap;`">{5}</td>
            <td style=`"$style`">{6}</td>
        </tr>
        " -f $Events[$i].MachineName.ToLower(),
                $Events[$i].ProviderName,
                $Events[$i].Id,
                $Events[$i].Count,
                $Events[$i].LevelDisplayName,
                $Events[$i].LogName,
                $message
    }

    $Body += @"
</table>
</body>
</html>
"@

    $Body = [String]::Join("`n", $Body)

    Write-Verbose "Sending Email"
    Send-MailMessage -Body $Body -BodyAsHtml -From $From -To $To -SmtpServer $SmtpServer -Subject $Subject -UseSsl
}
