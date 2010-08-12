################################################################################
# 
# $URL$
# $Author$
# $Date$
# $Rev$
# 
# DESCRIPTION:  General functions added to the environment.
#
# 
# Copyright (c) 2009,2010 Seth Wright <wrightst@jmu.edu>
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
#
################################################################################

# If the function already exists in this runspace, remove it so it 
# can be re-added below.
if (Test-Path function:Get-HexDump) { 
    Remove-Item function:Get-HexDump
}

function global:Get-HexDump($Path,$Width=10, $Bytes=-1)
{
    $OFS=""
    Get-Content -Encoding byte $Path -ReadCount $Width `
    -totalcount $Bytes | Foreach-Object {
        $byte = $_
        if (($byte -eq 0).count -ne $Width)
        {
            $hex = $byte | Foreach-Object {
                " " + ("{0:x}" -f $_).PadLeft(2,"0")}
            $char = $byte | Foreach-Object {
                if ([char]::IsLetterOrDigit($_))
                { [char] $_ } else { "." }}
            "$hex $char"
        }
    }
}

Write-Host "`tAdded Get-HexDump to global functions." -Fore White


if (Test-Path function:ConvertTo-Base64) { 
    Remove-Item function:ConvertTo-Base64
}
function global:ConvertTo-Base64($string) {
   $bytes  = [System.Text.Encoding]::UTF8.GetBytes($string);
   $encoded = [System.Convert]::ToBase64String($bytes); 

   return $encoded;
}

Write-Host "`tAdded ConvertTo-Base64 to global functions." -Fore White


if (Test-Path function:ConvertFrom-Base64) { 
    Remove-Item function:ConvertFrom-Base64
}
function global:ConvertFrom-Base64($string) {
   $bytes  = [System.Convert]::FromBase64String($string);
   $decoded = [System.Text.Encoding]::UTF8.GetString($bytes); 

   return $decoded;
}

Write-Host "`tAdded ConvertFrom-Base64 to global functions." -Fore White
