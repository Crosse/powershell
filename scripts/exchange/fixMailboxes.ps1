param ([switch]$Force=$false)

if ($mailboxesToCheck -eq $null) {
    Write-Host "var `$mailboxesToCheck does not exist.  Getting all Mailboxes to check."
    $mailboxesToCheck = Get-Mailbox -ResultSize Unlimited
}

if ($mailboxesToCheck -eq $null) {
    Write-Host "No mailboxes to fix!"
    return
}

$count = $mailboxesToCheck.Count
if ($count -eq 0 -or $count -eq $null) { $count = 1 }
$i = 0
foreach ($userToCheck in $mailboxesToCheck) {
    $percent = $([int]($i/$count*100))
    Write-Progress -Activity "Checking Mailboxes for invalid addresses" `
        -Status "$percent% Complete" `
        -PercentComplete $percent -CurrentOperation "Verifying..."
    $i++

    if ($userToCheck -eq $null) {
        Write-Error "Something went horribly wrong.  I'm stopping now."
        return
    }

    # Get a fresh copy of the data.
    $user = Get-Mailbox $userToCheck.Identity
    if ($user -eq $null) { 
        Write-Warning "$($userToCheck.Name): Could not find in Active Directory"
        continue
    }

    $emailAddresses = $user.EmailAddresses

    $remove = New-Object System.Collections.ArrayList
    $emailAddresses | % {
        if ($_.PrefixString.ToLower() -notmatch "smtp") {
            continue
        }

        if (!$_.ProxyAddressString.ToLower().Contains("jmu.edu")) {
            $null = $remove.Add($_)
        } elseif ($_.ProxyAddressString.ToLower().Contains("notatthisaddress")) {
            $null = $remove.Add($_)
        } elseif ($_.ProxyAddressString.ToLower().Contains("dukes.jmu.edu")) {
            $null = $remove.Add($_)
        }
    }

    if ($remove.Count -gt 0) {
        Write-Host "$($userToCheck.Name): starting."
        $remove | % { 
            Write-Host "$($user.Name):  Removing old address $($_.ProxyAddressString)"
            $null = $emailAddresses.Remove($_)
        }

        if (!$Force) {
            Write-Host "Not setting attributes; use -Force to commit changes"
        } else {
            Set-Mailbox $user.Identity -EmailAddressPolicyEnabled:$false -EmailAddresses $emailAddresses
            Set-Mailbox $user.Identity -EmailAddressPolicyEnabled:$true
        }
        Write-Host "$($user.Name): done.`n"
    }
}

