[CmdletBinding()]
param (
    [Parameter(Mandatory=$false)]
    [ValidatePattern("[a-zA-Z]{2}")]
    [string]
    $StartingPrefix = "aa",

    [Parameter(Mandatory=$false)]
    [string]
    $StartingMailbox = ""
)

$a = [int][char]'a'
$z = [int][char]'z'

if (![String]::IsNullOrEmpty($StartingMailbox)) {
    $StartingPrefix = $StartingMailbox.Substring(0,2)
}

$totalPrefixes = [Math]::Pow(($z-$a+1), 2)
$processedPrefixes = 0
$processedMailboxes = 0

foreach ($1 in $a..$z) {
    foreach ($2 in $a..$z) {
        $prefix = ("{0}{1}" -f [char]$1, [char]$2)

        $complete = [Int32]($processedPrefixes/$totalPrefixes * 100)
        Write-Progress -Id 1 -Activity "Processing mailboxes" -Status "Progress ($processedPrefixes of $totalPrefixes):" `
                       -CurrentOperation "Prefix '$prefix'" -PercentComplete $complete

        if ($prefix -lt $StartingPrefix) {
            #Write-Verbose "Skipping $prefix (comes before requested starting prefix of $StartingPrefix)"
            $processedPrefixes++
            continue
        }

        Write-Verbose "Getting all mailboxes for prefix '$prefix'"
        $mboxes = Get-Mailbox -ResultSize Unlimited $prefix* -ErrorAction SilentlyContinue | Sort-Object Name
        $processedMailboxes = 0
        foreach ($mbox in $mboxes) {
            if ($mbox -eq $null) { continue }

            $pctComplete = if ($mboxes.Count -eq 0) { 0 } else { [Int32]($processedMailboxes/$mboxes.Count * 100) }
            Write-Progress -Id 2 -ParentId 1 -Activity "Processing mailboxes starting with $prefix" `
                           -Status "Progress ($processedMailboxes of $($mboxes.Count)):" -CurrentOperation $mbox.Name -PercentComplete $pctComplete

            if (![String]::IsNullOrEmpty($StartingMailbox) -and $mbox.Name -lt $StartingMailbox) {
                Write-Verbose "Skipping $($mbox.Name) (comes before requested starting mailbox $StartingMailbox)"
                $processedMailboxes++
                continue
            }
            $StartingMailbox = ""

            #Write-Verbose "Working on $($mbox.Name)"
            $rc = Get-MailboxRegionalConfiguration $mbox.Identity

            $redirect = $mbox.ForwardingSmtpAddress
            if ([String]::IsNullOrEmpty($mbox.ForwardingSmtpAddress)) {
                $ir = @(Get-InboxRule -Mailbox $mbox.Identity | ? { $_.RedirectTo -ne $null -and $_.Enabled })
                if ($ir.Count -gt 0) {
                    $redirect = $ir[0].RedirectTo[0].Address
                    if ($redirect.StartsWith("/o=")) {
                        $redirect = (Get-Recipient $redirect).PrimarySmtpAddress.ToString()
                    }
                }
            }

            New-Object PSObject -Property @{
                Name = $mbox.Name
                PrimarySmtpAddress = $mbox.PrimarySmtpAddress
                RedirectTo = $redirect
                MailboxAccessed = if ($rc.Language) { $true } else { $false }
            }
            $processedMailboxes++
        }
        Write-Progress -Id 2 -ParentId 1 -Activity "Processing mailboxes starting with $prefix" -Status "Progress:" -Completed
        $processedPrefixes++

    }
}
Write-Progress -Id 1 -Activity "Processing mailboxes" -Status "Progress:" -Completed
