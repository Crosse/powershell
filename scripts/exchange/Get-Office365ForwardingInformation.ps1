[CmdletBinding()]
param (
    [Parameter(Mandatory=$false)]
    [ValidatePattern("[a-zA-Z]{2}")]
    [string]
    $StartingPrefix = "aa",

    [Parameter(Mandatory=$false)]
    [string]
    $StartingMailbox = "",

    [Parameter(Mandatory=$true)]
    [System.Management.Automation.PSCredential]
    $Credential,

    [Parameter(Mandatory=$false)]
    [System.Management.Automation.Runspaces.PSSession]
    $Session = $null,

    [Parameter(Mandatory=$false)]
    [int]
    $BatchSize = 1000
)

function New-Office365Session {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory=$true)]
        [System.Management.Automation.PSCredential]
        $Credential
    )

    Write-Verbose "Destroying any previous sessions"
    Get-PSSession | ? { $_.ComputerName -match 'outlook.com' -and $_.ConfigurationName -eq 'Microsoft.Exchange' } | Remove-PSSession

    $msoExchangeURL = "https://outlook.office365.com/powershell-liveid"
    Write-Verbose "Creating new session to $msoExchangeURL"
    $oldPref = $VerbosePreference
    $VerbosePreference = "SilentlyContinue"
    $session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri $msoExchangeURL -ErrorAction Stop `
                             -Credential $Credential -Authentication Basic -AllowRedirection -Verbose:$false
    Write-Verbose "Importing session"
    $null = Import-PSSession $session -Prefix Office365 -AllowClobber -Verbose:$false
    $VerbosePreference = $oldPref
    return $session
}



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

        if ($Session -eq $null -or $Session.State -ne "Opened") {
            $Session = New-Office365Session -Credential $Credential
        }

        Write-Verbose "Getting all mailboxes for prefix '$prefix'"
        $mboxes = Get-Office365Mailbox -ResultSize Unlimited $prefix* -ErrorAction SilentlyContinue | Sort-Object Name

        $processedMailboxes = 0
        foreach ($mbox in $mboxes) {
            if ($mbox -eq $null) { continue }

            $pctComplete = [Int32]($processedMailboxes/$mboxes.Count * 100)
            Write-Progress -Id 2 -ParentId 1 -Activity "Processing mailboxes starting with $prefix" `
                           -Status "Progress ($processedMailboxes of $($mboxes.Count)):" -CurrentOperation $mbox.Name -PercentComplete $pctComplete

            if (![String]::IsNullOrEmpty($StartingMailbox) -and $mbox.Name -lt $StartingMailbox) {
                Write-Verbose "Skipping $($mbox.Name) (comes before requested starting mailbox $StartingMailbox)"
                $processedMailboxes++
                continue
            }
            $StartingMailbox = ""

            if ($Session -eq $null -or $Session.State -ne "Opened") {
                $Session = New-Office365Session -Credential $Credential
            }

            #Write-Verbose "Working on $($mbox.Name)"
            $rc = Get-Office365MailboxRegionalConfiguration $mbox.Identity

            $redirect = $mbox.ForwardingSmtpAddress
            if ([String]::IsNullOrEmpty($mbox.ForwardingSmtpAddress)) {
                $ir = @(Get-Office365InboxRule -Mailbox $mbox.Identity | ? { $_.RedirectTo -ne $null })
                $redirect = $ir[0].RedirectTo
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
