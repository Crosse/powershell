[CmdletBinding(SupportsShouldProcess=$true,
        ConfirmImpact="High")]
param (
        [string]
        $From = "Exchange System <it-exmaint@jmu.edu>",

        [string]
        $SmtpServer = "mailgw.jmu.edu",

        [string]
        $Path = $PWD
      )

$processedPath = Join-Path $Path "Processed"
$files = Get-ChildItem (Join-Path $Path "*.csv")

if ($files -eq $null) {
    Send-MailMessage -From $From -To "wrightst@jmu.edu" -SmtpServer $SmtpServer `
        -UseSSL -Subject "New Mailbox Notifications:  Nothing to do!" -Body "No files to process."
    return
}

$csv = $files | % { Import-Csv $_ } | Sort Address
Write-Verbose "Found $($csv.Count) items"

try {
    if ((Test-Path $processedPath -PathType Container) -eq $false) {
        New-Item -ItemType Directory $processedPath -ErrorAction Stop
    }
    Move-Item $files $processedPath -ErrorAction Stop
} catch {
    $e = $_
    Send-MailMessage -From $From -To "wrightst@jmu.edu" -SmtpServer $SmtpServer `
        -UseSSL -Subject "New Mailbox Notifications:  FAILED" -Body $e
    throw $e
}

$encoder = New-Object System.Text.ASCIIEncoding
$rejectedItems = @()
$badItems = @()

foreach ($item in $csv) {
    $To = $item.Address
    $Subject = $item.Subject
    $Body = $encoder.GetString([Convert]::FromBase64String($item.MessageBody))

    if ([String]::IsNullOrEmpty($From) -or
            [String]::IsNullOrEmpty($To) -or
            [String]::IsNullOrEmpty($Subject) -or
            [String]::IsNullOrEmpty($Body)) {
        Write-Error "One or more fields was empty!"
        $badItems += $item
    } else {
        $desc = "Send Email to $To"
        $caption = $desc
        $warning = "Are you sure you want to perform this action?`n"
        $warning += "This will send a notification email to $To with the subject: "
        $warning += "`"$Subject`""

        if (!$PSCmdlet.ShouldProcess($desc, $warning, $caption)) {
            continue
        } else {
            Write-Verbose "Sending email to '$To' from '$From' with subject '$Subject' and body with length of $($Body.Length)"
            $error.Clear()
            try {
                Send-MailMessage `
                    -From $From -To $To -Bcc $From `
                    -Subject $Subject -Body $Body -BodyAsHtml `
                    -SmtpServer $SmtpServer -UseSsl -ErrorAction Stop
            } catch {
                Write-Error $_
                Write-Verbose "Recording rejection in rejected.log"
                $rejectedItems += $item
            }
        }
    }
}

$rejectedItems | ConvertTo-Csv | Out-File -Append -Encoding ASCII (Join-Path $Path "rejected.log")
$badItems | ConvertTo-Csv | Out-File -Append -Encoding ASCII (Join-Path $Path "bad.log")
