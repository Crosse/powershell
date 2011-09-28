param ([switch]$Force=$false)

if ($mailUsersToCheck -eq $null) {
    Write-Host "var `$mailUsersToCheck does not exist.  Getting all MailUsers that must be fixed."
    $mailUsersToCheck = Get-MailUser -ResultSize Unlimited -Filter { ExternalEmailAddress -like '*@dukes.jmu.edu' -and WindowsEmailAddress -notlike '*@dukes.jmu.edu' }
}

if ($mailUsersToCheck -eq $null) {
    Write-Host "No users to fix!"
    return
}

$count = $mailUsersToCheck.Count
if ($count -eq 0 -or $count -eq $null) { $count = 1 }

Write-Host "Found $count users to fix"

$i = 0
foreach ($userToCheck in $mailUsersToCheck) {
    $percent = $([int]($i/$count*100))
    Write-Progress -Activity "Fixing MailUsers with invalid addresses" `
        -Status "$percent% Complete" `
        -PercentComplete $percent -CurrentOperation "$userToCheck"
    $i++
    if ($userToCheck -eq $null) {
        Write-Error "Something went horribly wrong.  I'm stopping now."
        return
    }

    Write-Host "$($userToCheck.Name): starting."

    # Get a fresh copy of the data.
    $user = Get-MailUser $userToCheck.Identity
    if ($user -eq $null) { 
        Write-Warning "$($userToCheck.Name): Could not find in Active Directory"
        continue
    }

    $newAddr = $user.ExternalEmailAddress.SmtpAddress
    $emailAddresses = $user.EmailAddresses

    if (!($emailAddresses.Contains($user.ExternalEmailAddress)) ) {
        Write-Host "$($user.Name):  Adding address `"$($user.ExternalEmailAddress.SmtpAddress)`""
        $null = $emailAddresses.Add($user.ExternalEmailAddress)
    }

    $remove = New-Object System.Collections.ArrayList
    $emailAddresses | % {
        $null = $_.ToSecondary()
        if (!$_.ProxyAddressString.ToLower().Contains("jmu.edu")) {
            $null = $remove.Add($_)
        } elseif ($_.ProxyAddressString.ToLower().Contains("notatthisaddress")) {
            $null = $remove.Add($_)
        }
    }

    if ($remove.Count -gt 0) {
        $remove | % { 
            Write-Host "$($user.Name):  Removing old address $($_.ProxyAddressString)"
            $null = $emailAddresses.Remove($_)
        }

        if (!$Force) {
            Write-Host "Not setting attributes; use -Force to commit changes"
        } else {
            Set-MailUser $user.Identity -EmailAddressPolicyEnabled:$false -WindowsEmailAddress $newAddr -EmailAddresses $emailAddresses
            Set-MailUser $user.Identity -EmailAddressPolicyEnabled:$true
        }
        Write-Host "$($user.Name): done.`n"
    }
}

