################################################################################
# 
# NAME  : 
# AUTHOR: Seth Wright , James Madison University
# DATE  : 5/13/2009
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


if (Test-Path function:ps_tail) { Remove-Item function:ps_tail }

# ps_tail file -lines
function global:ps_tail {
  param ($path, $lines="-10", [switch] $f)

  if ($lines.GetType().Name -eq "String" ) {
    $lines = [string]$lines.trim("-")
  }
  if ($path -ne $Null) {
    if (Test-Path $path) {
      $content = Get-Content -path $path
      for($i=$content.length - $lines; $i -le $content.length; $i++) {
        $content[$i]
      }
    }
  }
}

Write-Host "Added ps_tail to global functions." -Fore White
