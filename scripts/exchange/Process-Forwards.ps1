param (
        [ValidateNotNullOrEmpty()]
        [ValidateScript( {Test-Path $_} )]
        $FilePath
      )

$forwards = Import-Csv -Path $FilePath

Write-Output "Total users:`t`t$($forwards.Count)"

$forwarders = New-Object System.Collections.ArrayList
$nonForwarders = New-Object System.Collections.ArrayList
$processed = 0

$forwards | % {
    if ($_.MailboxAccessed -eq "True") {
        $processed += 1
        if ($_.RedirectTo -ne "") {
            $null = $forwarders.Add($_)
        } else {
            $null = $nonForwarders.Add($_)
        }
    }
}

Write-Output "Total activated users:`t$processed"
Write-Output "Users forwarding:`t$($forwarders.Count) ($([Math]::Round($forwarders.Count / $processed * 100))%)"
Write-Output "Users NOT forwarding:`t$($nonForwarders.Count) ($([Math]::Round($nonForwarders.Count / $processed * 100))%)"

Write-Output "`nBreakdown of forwarding domains:"

$forwarders | % {
    if ($_.RedirectTo -match "SMTP") {
        ([System.Net.Mail.MailAddress]($_.RedirectTo.Substring($_.RedirectTo.IndexOf(":") + 1).Replace("]", ""))).Host
    }
} | Group -NoElement | Sort Count -Descending | Select -First 10
