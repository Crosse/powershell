param ([string]$User=$null, $Credential=$(Get-Credential))

if ([String]::IsNullOrEmpty($User)) {
    Write-Error "No user given."
    return
}

#$DomainController = (gc Env:\LOGONSERVER).Replace('\', '')
$DomainController = "cisatgc.cisat.jmu.edu"

if ($DomainController -eq $null) { 
    Write-Warning "Could not determine the local computer's logon server!"
    return
}


$objUser = Get-User -ErrorAction SilentlyContinue -DomainController $DomainController -Credential $Credential -Identity $User

if ($objUser -eq $null -or $objUser.RecipientTypeDetails -notmatch 'UserMailbox') {
    Write-Error "Could not find user in AD."
    return
}

$objUser

$objUser | Disable-Mailbox -DomainController $DomainController
$objUser | Enable-MailUser -DomainController $DomainController -ExternalEmailAddress $User@ad.jmu.edu -Confirm:$true
