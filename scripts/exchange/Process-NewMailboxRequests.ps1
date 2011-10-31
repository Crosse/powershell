param (
        [Parameter(Mandatory=$true)]
        [ValidateScript({Test-Path $_})]
        # Where to find the CSV files to process.
        [string]
        $FilePath,

        [Parameter(Mandatory=$true)]
        [ValidateScript({(Test-Path $_) -and ($_ -ne $FilePath)})]
        # Where to put the processed files.
        [string]
        $ProcessedPath,

        [Parameter(Mandatory=$false)]
        [switch]
        $SendEmail=$false,

        [Parameter(Mandatory=$false)]
        [ValidateScript({
            if ($SendEmail -and [System.String]::IsNullOrEmpty($_)) {
                return $false
            } else {
                return $true
            }})]
        [string]
        $From,

        [Parameter(Mandatory=$false)]
        [ValidateScript({
            if ($SendEmail -and [System.String]::IsNullOrEmpty($_)) {
                return $false
            } else {
                return $true
            }})]
        [string[]]
        $To,

        [Parameter(Mandatory=$false)]
        [ValidateNotNull()]
        # The SMTP server to use.
        [string]
        $SmtpServer
      )

$module = Import-Module -Force -PassThru .\UserProvisioning.psm1
if ($module -eq $null) {
    $Subject = "Mailbox Provisioning:  Could not load UserProvisioning Module!"
    $output = "Could not import module UserProvisioning.psm1."
    Send-MailMessage -From $From -To $To -Subject $Subject -Body $output -SmtpServer $SmtpServer
    return
}

$module = Import-Module -Force -PassThru "C:\Program Files\Common Files\Microsoft Lync Server 2010\Modules\Lync\Lync.psd1"
if ($module -eq $null) {
    $Subject = "Mailbox Provisioning:  Could not load Lync Module!"
    $output = "Could not import module Lync.psd1."
    Send-MailMessage -From $From -To $To -Subject $Subject -Body $output -SmtpServer $SmtpServer
    return
}

$files = Get-ChildItem (Join-Path $FilePath "*.csv")
if ($files -eq $null) {
    $Subject = "Mailbox Provisioning:  Nothing to do!"
    $output = "No files to process."
} else {
    $users = $files | Import-Csv -Header User,Date,Reason | Sort User -Unique
    if ($users -eq $null) {
        $Subject = "Mailbox Provisioning: Nothing to do!"
        $output = "Files existed, but were empty."
    } else {
        $errorCount = 0
        $output = ""
        foreach ($user in $users) {
            if ($user.User -eq "User") {
                # This is a header line, skip it.
                $user.Reason = $null
                continue
            }

            try {
                $user.Date = Get-Date -Format o $user.Date
            } catch {
                $user.Date = Get-Date -Format o
            }
            $result = Add-ProvisionedMailbox `
                        -Identity $user.User `
                        -MailboxLocation Local `
                        -EnableForLync `
                        -LyncRegistrarPool lyncpool.jmu.edu `
                        -MailContactOrganizationalUnit 'ad.jmu.edu/ExchangeObjects/MailContacts' `
                        -SendEmailNotification:$false `
                        -Confirm:$false

            if ($result.ProvisioningSuccessful -eq $false) {
                $user.Reason = $result.Error
                $errorCount++
                $output += "FAILURE: [ {0,-8} ] - {1}`n" -f $user.User, $result.Error
            } else {
                $user.Reason = $null
                $output += "SUCCESS: [ {0,-8} ] - {1}" -f $user.User, $result.Error
                if ($result.EnabledForLync -and $result.MailContactCreated) {
                    $output += " (Lync Enabled / MailContact Created)"
                } elseif ($result.EnabledForLync) {
                    $output += " (Lync Enabled)"
                } elseif ($result.MailContactCreated) {
                    $output += " (MailContact Created)"
                }
                $output += "`n"
            }
        }

        $erroredOut = $users | ? { $_.Reason -ne $null } | Sort User -Unique
        if ($erroredOut -ne $null) {
            $erroredOut |
                Export-Csv -NoTypeInformation `
                    -Encoding ASCII `
                    -Path (Join-Path $FilePath "errors_$(Get-Date -Format yyyy-MM-dd_HH-mm-ss).csv")
        }

        if ($errorCount -gt 0) {
            $Subject = "Mailbox Provisioning: $errorCount errors detected"
        } else {
            $Subject = "Mailbox Provisioning: No errors detected"
        }
    }

    Move-Item -Force $files $ProcessedPath
}

# Yes, I know this is ugly.  I don't care, because it works.
$output = $output.Split("`n") | Sort -Unique
$sortedOutput = [System.String]::Join("`n", $output)

Write-Host $Subject
Write-Host $sortedOutput

if ($SendEmail -eq $true) {
    Send-MailMessage -From $From -To $To -Subject $Subject -Body $sortedOutput -SmtpServer $SmtpServer
}
