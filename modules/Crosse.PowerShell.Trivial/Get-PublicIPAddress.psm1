################################################################################
#
# Copyright (c) 2016 Seth Wright <seth@crosse.org>
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

<#
    .SYNOPSIS
    Gets the computer's public IP address.

    .DESCRIPTION
    Gets the current public IP address of this computer by using "ipify", a simple IP address API (https://www.ipify.org/)

    .INPUTS
    None.  You cannot pipe data into this cmdlet.

    .OUTPUTS
    System.String.  Get-PublicIPAddress returns the public IP address.

    .EXAMPLE
    PS C:\> Get-PublicIPAddress
    203.0.113.147
#>
function Get-PublicIPAddress {
    [CmdletBinding()]
    param ()

    Invoke-RestMethod -UseBasicParsing 'https://api.ipify.org'
}
