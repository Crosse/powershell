################################################################################
# 
# $URL$
# $Author$
# $Date$
# $Rev$
# 
# Copyright (c) 2009 Seth Wright (wrightst@jmu.edu)
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

$now = Get-Date
$fileName = "SafeListUpdate_$($now.Year)-$($now.Month)-$($now.Day).log"
Start-Transcript -Path $fileName

$mailboxes = Get-Mailbox -ResultSize Unlimited
$i = 0
$j = 0

foreach ($mailbox in $mailboxes) {
    Update-SafeList -Identity $mailbox
    if ($?) {
        $j++
    } else {
        Write-Host "SafeList update was unsuccessful for $($mailbox.Alias)"
    }
    $i++
    Write-Progress -Activity "Updating Safelists..." -Status "$($mailbox.Alias)" -percentComplete ($i/$mailboxes.Count*100)
}

Write-Host "Sucessfully processed SafeList data for $($j)/$($i) mailboxes."

Stop-Transcript
