################################################################################
# 
# $Id$
# 
# DESCRIPTION:  Sends an email with relevant Edge statistics to various
#               users.
#
# 
# Copyright (c) 2009,2010 Seth Wright <wrightst@jmu.edu>
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

if (!(Get-PSSnapin -Name FSSPSSnapIn -ErrorAction SilentlyContinue)) {
    Add-PSSnapin FSSPSSnapIn
}

if (!(Get-PSSnapin -Name Microsoft.Exchange.Management.PowerShell.Admin -ErrorAction SilentlyContinue)) {
    Add-PSSnapin Microsoft.Exchange.Management.PowerShell.Admin
}

$startTime = $(Get-Date).AddDays(-1)
$endTime = $(Get-Date)

$localhost = (Get-ChildItem Env:\COMPUTERNAME).Value
$From = "Exchange System <it-exmaint@jmu.edu>"
$To = "wrightst@jmu.edu"
$Title = "Edge Detail for $localhost - $(Get-Date -Format d)"

$log = Get-MessageTrackingLog -ResultSize Unlimited -Start $startTime
    
$summary = $log | Group-Object EventId,ConnectorId -NoElement | `
                Sort Count -Desc | ft -AutoSize | Out-String

$topExtDomains = $log | where { $_.EventId -eq 'RECEIVE' -and `
                $_.ClientIp -notmatch '134.126.52.\d\d\d' } | `
                Group-Object -Property { $_.Sender.ToString().Split('@')[1] } `
                -NoElement | Sort Count -Desc | Select-Object -First 10 | `
                ft -AutoSize | Out-String

#$topIntDomains = $log | where { $_.EventId -eq 'RECEIVE' -and `
#                $_.ClientIp -match '134.126.52.\d\d\d' } | `
#                Group-Object -Property { $_.Sender.ToString().Split('@')[1] } `
#                -NoElement | Sort Count -Desc | Select-Object -First 10 | `
#                ft -AutoSize | Out-String

Remove-Variable log


#$spamReport = Get-FseSpamReport -Starttime $startTime -Endtime $endTime | Out-String

$spamLog = Get-FseSpamAgentLog -After $startTime -Before $endTime | `
            group Action,Reason,ReasonData -NoElement | Sort Count -Desc | `
            ft -AutoSize | Out-String

$ffHealth = Get-FseHealth | where { $_.Status -ne "GREEN" } | select LastUpdate,Message | ft -autosize | Out-String
if ($ffHealth.Length -eq 0) {
    $ffHealth = "Green across the board!" 
}

$updateStatus = Get-FseSignatureUpdate | where { $_.UpdateStatus -ne 'EngineUpdateNotAttempted' } | Out-String
if ($updateStatus.Length -eq 0) {
    $updateStatus = "All engines are up-to-date" 
}

$ffReport = Get-FseReport -Type AllFilters | where { $_.ScanJob -match 'Transport' } | Out-String

$Body  = @"
#######################################################################
# Mail Totals by EventId and Connector                                #
#######################################################################
$summary

#######################################################################
# Forefront Spam Agent Log Summary                                    #
#######################################################################
$spamLog

#######################################################################
# Top 10 Sender Domains (Inbound)                                     #
#######################################################################
$topExtDomains

#######################################################################
# Forefront Signature Update Status                                   #
#######################################################################
$updateStatus

#######################################################################
# Forefront Health Summary                                            #
#######################################################################
$ffHealth

#######################################################################
# Forefront Report Summary (since last counter clear)                 #
#######################################################################
$ffReport
"@

$SmtpServer = $localhost
$SmtpClient = New-Object System.Net.Mail.SmtpClient
$SmtpClient.Credentials = [System.Net.CredentialCache]::DefaultNetworkCredentials
$SmtpClient.Port = 25
$SmtpClient.host = $SmtpServer
$SmtpClient.Send($From, $To, $Title, $Body)

