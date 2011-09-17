[CmdletBinding(SupportsShouldProcess=$true,
        ConfirmImpact="High")]

param (
        [Parameter(Mandatory=$true,
            ValueFromPipelineByPropertyName=$true)]
        [Alias("LegacyDN")]
        [string]
        [ValidateNotNullOrEmpty()]
        # The identity of the mailbox that should have its quota modified.
        $Identity,

        [UInt32]
        [ValidateRange(0, 2GB)]
        # The amount of headroom to give the mailbox.  Default is 100MB.
        # This value is also set as the amount of space between IsssWarning
        # and ProhibitSend, and between ProhibitSend and ProhibitSendReceive.
        $Headroom=100MB,

        [Int64]
        # The upper limit above which no mailbox quota should be raised.
        $UpperLimit=3GB,

        [Parameter(Mandatory=$false)]
        [string]
        # The domain controller to use for all operations.
        $DomainController
      )

# This section executes only once, before the pipeline.
BEGIN {
    Write-Verbose "Performing initialization actions."

    if ([String]::IsNullOrEmpty($DomainController)) {
        $dc = [System.DirectoryServices.ActiveDirectory.Domain]::`
                GetCurrentDomain().FindDomainController().Name

        if ($dc -eq $null) {
            Write-Error "Could not find a domain controller to use for the operation."
            return
        }
    } else {
        $dc = $DomainController
    }

    Write-Verbose "Using Domain Controller $dc"
    Write-Verbose "Initialization complete."
} # end 'BEGIN{}'


# This section executes for each object in the pipeline.
PROCESS {
    $Mailbox = Get-Mailbox $Identity -DomainController $dc
    if ($Mailbox -eq $null) {
        Write-Error "Could not find mailbox."
        return
    }

    $stats = Get-MailboxStatistics $Mailbox -DomainController $dc
    if ($stats -eq $null) {
        Write-Error "Could not get mailbox statistics"
        return
    }

    if ($stats.StorageLimitStatus -ne 'NoChecking' -and 
            $stats.StorageLimitStatus -ne 'BelowLimit') {
        if ($Mailbox.UseDatabaseQuotaDefaults -eq $true) {
            $db = Get-MailboxDatabase $Mailbox.Database `
                        -DomainController $dc

            Write-Verbose "Mailbox is using database quota defaults"
            $IssueWarning           = $db.IssueWarningQuota.Value.ToMB()
            $ProhibitSend           = $db.ProhibitSendQuota.Value.ToMB()
            $ProhibitSendReceive    = $db.ProhibitSendReceiveQuota.Value.ToMB()
        } else {
            Write-Verbose "Mailbox has custom quota levels"
            $IssueWarning           = $Mailbox.IssueWarningQuota.Value.ToMB()
            $ProhibitSend           = $Mailbox.ProhibitSendQuota.Value.ToMB()
            $ProhibitSendReceive    = $Mailbox.ProhibitSendReceiveQuota.Value.ToMB()
        }

        Write-Verbose "Current Quota Limits for mailbox:"
        Write-Verbose "IssueWarningQuota:         $($IssueWarning)MB"
        Write-Verbose "ProhibitSendQuota:         $($ProhibitSend)MB"
        Write-Verbose "ProhibitSendReceiveQuota:  $($ProhibitSendReceive)MB"

        $totalItemSize = [Math]::Floor($stats.TotalItemSize.Value.ToBytes()/1MB)
        Write-Verbose "TotalItemSize:             $($totalItemSize)MB"

        # Round up to the nearest 100MB.
        $totalRoundedSize = ($totalItemSize - (($totalItemSize % 100) - 100))*1MB
        Write-Verbose "Total Rounded Size:        $($totalRoundedSize/1MB)MB"

        $IssueWarning           = $totalRoundedSize + $Headroom
        $ProhibitSend           = $totalRoundedSize + $Headroom * 2
        $ProhibitSendReceive    = $totalRoundedSize + $Headroom * 3

        if ($IssueWarning -gt $UpperLimit) {
            Write-Error "Raising the user's quota would exceed the requested upper limit."
            return
        }

        Write-Verbose "Proposed Quota Limits for mailbox:"
        Write-Verbose "IssueWarningQuota:         $($IssueWarning/1MB)MB"
        Write-Verbose "ProhibitSendQuota:         $($ProhibitSend/1MB)MB"
        Write-Verbose "ProhibitSendReceiveQuota:  $($ProhibitSendReceive/1MB)MB"

        $desc = "Raise Quota for `"$Mailbox`""
        $caption = $desc
        $warning = "Are you sure you want to perform this action?`n"
        $warning += "This will raise quota limits for the user "
        $warning += "to $($IssueWarning/1MB)MB/$($ProhibitSend/1MB)MB/$($ProhibitSendReceive/1MB)MB.`n"
        $warning += "Current mailbox size:  $($TotalItemSize)MB"

        if (!$PSCmdlet.ShouldProcess($desc, $warning, $caption)) {
            return
        }

        Set-Mailbox -Identity $Mailbox -DomainController $dc `
                    -IssueWarningQuota $IssueWarning `
                    -ProhibitSendQuota $ProhibitSend `
                    -ProhibitSendReceiveQuota $ProhibitSendReceive `
                    -UseDatabaseQuotaDefaults:$false
    } else {
        Write-Verbose "Mailbox is within quota limits"
    }

    Get-Mailbox -Identity $Mailbox -DomainController $dc
}
