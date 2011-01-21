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
        $ProcessedPath
      )

Import-Module ..\modules\Crosse.PowerShell.Exchange\UserProvisioning.psm1

$files = Get-ChildItem (Join-Path $FilePath "*.csv")

if ($files -eq $null) {
    $Subject = "Mailbox Provisioning:  Nothing to do!"
    $output = "No files to process."
} else {
    $users = $files | Import-Csv -Header User,Date,Reason
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
                        -MailContactOrganizationalUnit 'ad.test.jmu.edu/ExchangeObjects/MailContacts' `
                        -Confirm:$false

            if ($result.ProvisioningSuccessful -eq $false) {
                $user.Reason = $result.Error
                $errorCount++
                $output += "FAILURE:  $($user.User):  $($user.Reason)`n"
            } else {
                $user.Reason = $null
                $output += "SUCCESS:  $($user.User):  $($result.MailboxLocation) mailbox provisioned"
                if ($result.MailContactCreated -eq $true) {
                    $output += " (MailContact created to preserve previous MailUser info)"
                }
                $output += ".`n"
            }
        }

        $users | ? { $_.Reason -ne $null } | 
            Export-Csv -NoTypeInformation `
                -Encoding ASCII `
                -Path (Join-Path $FilePath "errors_$(Get-Date -Format yyyy-MM-dd_HH-mm-ss).csv")

        if ($errorCount -gt 0) {
            $Subject = "Mailbox Provisioning: $errorCount errors detected"
        } else {
            $Subject = "Mailbox Provisioning: No errors detected"
        }
    }

    Move-Item $files $ProcessedPath
}

Write-Host $Subject
Write-Host $output
Send-MailMessage -From it-exmaint@test.jmu.edu -To "wrightst@jmu.edu" -Subject $Subject -Body $output -SmtpServer exchangetest.jmu.edu


