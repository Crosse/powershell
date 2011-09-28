Write-Output ""
Write-Output "Provide credentials for Dukes or DukesDev"
Write-Output ""
$LiveCred = Get-Credential
Write-Output ""
Write-Output "Connect to Outlook.com"
Write-Output ""
$Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://pod51004psh.outlook.com/PowerShell -Credential $LiveCred -Authentication Basic -AllowRedirection
Write-Output ""
Write-Output "Import Session"
Write-Output ""
Import-PSSession $Session -AllowClobber
