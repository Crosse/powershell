################################################################################
# 
# NAME  : Get-PhishingReplyAddresses.ps1
# AUTHOR: Seth Wright , James Madison University
# DATE  : 6/12/2009
# 
# Copyright (c) 2009 Seth Wright
#
# Permission to use, copy, modify, and distribute this software for any
# purpose with or without fee is hereby granted, provided that the above
# copyright notice and this permission notice appear in all copies.
#
# THE SOFTWARE IS PROVIDED "AS IS" AND THE AUTHOR DISCLAIMS ALL WARRANTIES
# WITH REGARD TO THIS SOFTWARE INCLUDING ALL IMPLIED WARRANTIES OF
# MERCHANTABILITY AND FITNESS. IN NO EVENT SHALL THE AUTHOR BE LIABLE FOR
# ANY SPECIAL, DIRECT, INDIRECT, OR CONSEQUENTIAL DAMAGES OR ANY DAMAGES
# WHATSOEVER RESULTING FROM LOSS OF USE, DATA OR PROFITS, WHETHER IN AN
# ACTION OF CONTRACT, NEGLIGENCE OR OTHER TORTIOUS ACTION, ARISING OUT OF
# OR IN CONNECTION WITH THE USE OR PERFORMANCE OF THIS SOFTWARE.
################################################################################

$wc = New-Object Net.WebClient

$now = Get-Date
$fileName = "PhishingList_$($now.Year)-$($now.Month)-$($now.Day).csv"

$wc.DownloadFile("http://anti-phishing-email-reply.googlecode.com/svn/trunk/phishing_reply_addresses", $fileName)

$phishAddresses = @()
foreach ($line in Get-Content $fileName) {
    if ($line.StartsWith("#")) {
        continue
    }
    $address = $line.Remove($line.IndexOf(','))
    $phishAddresses += $address
}

$Condition = Get-TransportRulePredicate AnyOfRecipientAddressContains
$Condition.Words = $phishAddresses

$Action1 = Get-TransportRuleAction PrependSubject
$Action1.Prefix = "[PHISHING RESPONSE] "
$Action2 = Get-TransportRuleAction RedirectMessage
$Action2.Addresses = @("postmaster@ad.jmu.edu")

$rule = Get-TransportRule 'PhishingAddresses' -ErrorAction SilentlyContinue
if ($rule) {
    Set-TransportRule -Conditions @(Condition) -Actions @($Action1, $Action2)
#    $rule.Conditions = @($Condition)
#    $rule.Actions = @($Action1, $Action2)
}
else {
    New-TransportRule -Name 'PhishingAddresses' -Condition @($Condition) -Action @($Action1, $Action2)
}
