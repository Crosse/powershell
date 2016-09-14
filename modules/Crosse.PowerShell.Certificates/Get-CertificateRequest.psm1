################################################################################
#
# Copyright (c) 2011-2016 Seth Wright <seth@crosse.org>
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

################################################################################
<#
    .SYNOPSIS
    Gets all pending certificate requests.

    .DESCRIPTION
    Gets pending certificate requests in either the User or Machine certificate
    store.

    .INPUTS
    None
        You cannot pipe objects to the cmdlet.

    .OUTPUTS
    System.String
        An indication of whether the cmdlet was successful is printed.

    .EXAMPLE
    C:\PS> Complete-CertificateRequest -CACertificateResponse .\response.cer -CertificateStore Machine
    Certificate Request completed successfully.
#>
################################################################################
function Get-CertificateRequest {
    [CmdletBinding()]
    param (
            [Parameter(Mandatory=$false)]
            [ValidateSet("Machine", "User", "All")]
            [string]
            # The certificate store in which to look for the pending certificate request.
            $CertifcateStore = "All"
          )

    if ($CertifcateStore -eq "All" -or $CertifcateStore -eq "Machine") {
        if (Test-Path "Cert:\LocalMachine\REQUEST") {
            Get-ChildItem Cert:\LocalMachine\REQUEST
        }
    }
    if ($CertifcateStore -eq "All" -or $CertifcateStore -eq "User") {
        if (Test-Path "Cert:\CurrentUser\REQUEST") {
            Get-ChildItem Cert:\CurrentUser\REQUEST
        }
    }
}