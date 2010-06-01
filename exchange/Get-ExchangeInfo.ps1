################################################################################
# 
# $Id$
# 
# DESCRIPTION:  Sends an email with relevant Exchange statistics to various
#               users.
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

# Change these to suit your environment
$SmtpServer = "it-exhub.ad.jmu.edu"
$From       = "it-exmaint@jmu.edu"
$To         = "wrightst@jmu.edu, gumgs@jmu.edu, liskeygn@jmu.edu, stockntl@jmu.edu, flynngn@jmu.edu, najdziav@jmu.edu"
#$To         = "wrightst@jmu.edu"
$Title      = "Exchange User Detail for $(Get-Date -Format d)"

##################################
$cwd = [System.IO.Path]::GetDirectoryName(($MyInvocation.MyCommand).Definition)

$dls = (Get-DistributionGroup).Count
$resources = 0
$shared = 0
$users = 0
$totalStorage = 0

$dbInfoArray = New-Object System.Collections.ArrayList
Get-MailboxServer | % { 
    $CMSName = $_.Name
    $OperationalMachines = (Get-ClusteredMailboxServerStatus -Identity $CMSName).OperationalMachines
    $null = $OperationalMachines | where { $_ -match "^(?<activenode>.*)\s+<Active.*" }
    
    if (!($matches.activenode)) {
        # No regex matches were found.
        $message = "Cannot determine the Active Node of CMS $CMSName"
        continue
    }

    $ActiveNode = $matches.activenode
    $preamble = "\\" + $ActiveNode + "\"

    $sgs = Get-StorageGroup -Server $CMSName
    $i = 1
    foreach ($sg in $sgs) {
        $percent = $([int]($i/$sgs.Count*100))
        Write-Progress -Activity "Processing Storage Groups on $CMSName" `
            -Status "$percent% Complete" `
            -PercentComplete $percent -CurrentOperation "Processing $($sg.Identity)"
        $i++
        
        foreach ($db in (Get-MailboxDatabase -StorageGroup $sg)) { 
            $dbInfo = New-Object PSObject
            $dbInfo = Add-Member -PassThru -InputObject $dbInfo NoteProperty Identity $null
            Add-Member -InputObject $dbInfo NoteProperty UserCount -1
            Add-Member -InputObject $dbInfo NoteProperty LogCount -1
            Add-Member -InputObject $dbInfo NoteProperty DbSize -1

            $dbInfo.Identity = $db.Identity

            $logPath = $preamble + $sg.LogFolderPath.ToString().Replace(":", "$")
            $dbInfo.LogCount = (gci "$logPath\*.log").Count

            $edbFilePath = $preamble + $db.EdbFilePath.ToString().Replace(":", "$")
            $dbSize = [Math]::Floor( (gci $edbFilePath).Length / 1MB )
            $totalStorage += $dbSize
            $dbInfo.DbSize = $dbSize

            $mboxes = Get-Mailbox -ResultSize Unlimited -Database $db
            $dbInfo.UserCount = $mboxes.Count

            $mboxes | % { 
                switch ($_.RecipientTypeDetails) {
                    'UserMailbox' { 
                        $users++
                        break
                    }
                    'SharedMailbox' { 
                        $shared++
                        break
                    }
                    'EquipmentMailbox' { 
                        $resources++
                        break
                    }
                    'RoomMailbox' {
                        $resources++
                        break
                    }
                    default { }
                }
            }
                        
            $null = $dbInfoArray.Add($dbInfo)
        }
    }
}

$recipientCounts = & "$cwd\Get-TopRecipientCounts.ps1"

$Body  = @"
     User Mailboxes:	$users
   Shared Mailboxes:	$shared
 Resource Mailboxes:	$resources
Distribution Groups:	$dls
 Total Storage Used:	$totalStorage MB
-----------------------------------------------------------------------

Top Senders by Total Recipient Count (last 24 hours):
$($recipientCounts | ft -Autosize | Out-String)
-----------------------------------------------------------------------
Database Information:
$($dbInfoArray | Sort Identity | ft -AutoSize | Out-String)
-----------------------------------------------------------------------
"@

$attachedFile = gci ([System.IO.Path]::GetTempFileName())
$attachedFile.MoveTo("$([System.IO.Path]::GetFileNameWithoutExtension($attachedFile.Name)).csv")

$dbInfoArray | Sort Identity | Export-Csv -NoTypeInformation -Encoding ASCII -Path $attachedFile

& "$cwd\Send-Email.ps1" -From $From -To $To -Subject $Title -Body $Body -AttachmentFile $attachedFile -SmtpServer $SmtpServer 

Remove-Item $attachedFile
