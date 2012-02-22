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
            # Gets only the events that correspond to the specified log
            # level(s).
            $Level
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
        if ([String]::IsNullOrEmpty($ComputerName) -eq $false) {
            Write-Verbose "Getting specified events for $ComputerName"
            $logs = Get-WinEvent -ComputerName $ComputerName -FilterHashTable $eventParams -Verbose:$false
        } else {
            Write-Verbose "Getting specified events for local machine"
            $logs = Get-WinEvent -FilterHashTable $eventParams -Verbose:$false
        }

        if ($logs -eq $null) {
            Write-Verbose "No events returned."
            return
        }

        $retval = @{}

        foreach ($event in $logs) {
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
        return ,@($retval.Values)
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

    $BodyStyle      = "font-family:sans-serif;color:black;background-color:white;"
    $baseStyle      = "border:1px solid green;border-collapse:collapse;padding:6px;"
    $tableStyle     = "$baseStyle;margin-left:auto;margin-right:auto;width:90%;"
    $captionStyle   = "border-style:none;border-collapse:collapse;text-align:center;font-weight:bold;font-size:1.2em;background-color:white;padding-left:6px;padding-right:6px;"
    $trStyle        = $baseStyle
    $trAltStyle     = "$trStyle;background-color:#EAF2D3;"
    $thStyle        = "$baseStyle;white-space:nowrap;text-align:center;color:#333333;"
    $tdStyle        = "$baseStyle;white-space:nowrap;"
    $tdEventIdStyle = $baseStyle
    $tdMessageStyle = $baseStyle

    $Body = @"
<!DOCTYPE html PUBLIC "-//W3C/DTD XHTML 1.0 Tranisional//EN"
"http://www.w3.org/xhtml1/DTD/xhtml1-transitional.dtd">
<html>
<head>
    <title>$Subject</title>
</head>
<body style="$BodyStyle">
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
            $style = $trStyle
        } else {
            $style = $trAltStyle
        }

        $Body += "
        <tr style=`"$style`">
            <td style=`"$tdStyle`">{0}</td>
            <td style=`"$tdEventIdStyle`">{1} {2}</td>
            <td style=`"$tdStyle`">{3}</td>
            <td style=`"$tdStyle`">{4}</td>
            <td style=`"$tdStyle`">{5}</td>
            <td style=`"$tdMessageStyle`">{6}</td>
        </tr>
        " -f $Events[$i].MachineName.ToLower(),
                $Events[$i].ProviderName,
                $Events[$i].Id,
                $Events[$i].Count,
                $Events[$i].LevelDisplayName,
                $Events[$i].LogName,
                [System.Web.HttpUtility]::HtmlEncode($Events[$i].SampleMessage.Split("`n")[0])
    }

    $Body += @"
</table>
</body>
</html>
"@

    $Body = [String]::Join("`n", $Body)

    Send-MailMessage -Body $Body -BodyAsHtml -From $From -To $To -SmtpServer $SmtpServer -Subject $Subject -UseSsl
}
