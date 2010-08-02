cd E:\IDM\ExchangeCSV

$files = gci *.csv
if ($files -eq $null) {
    $Subject = "Mailbox Provisioning:  Nothing to do!"
    $output = "No files to process."
} else {
    $users = gc $files | ConvertFrom-Csv -Header User,Date,Reason | % { $_.User } 

    if ($users -eq $null) {
        $Subject = "Mailbox Provisioning: Nothing to do!"
        $output = "Files existed, but were empty."
    } else {
        $output = $users | E:\Scripts\Provision-User.ps1 -Automated:$true -DomainController jmuadc4.ad.jmu.edu

        if ($? -eq $false) {
            $Subject = "Mailbox Provisioning: errors detected"
        } else {
            $Subject = "Mailbox Provisioning: No errors detected"
        }
    }

    move $files E:\IDM\ExchangeCSV\Processed
}

E:\Scripts\Send-Email.ps1 -From it-exmaint@jmu.edu -To "wrightst@jmu.edu,stockntl@jmu.edu,najdziav@jmu.edu" -Subject $Subject -Body $($output | Out-String) -SmtpServer it-exhub.ad.jmu.edu

