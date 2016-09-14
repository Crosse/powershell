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
    Completes a pending certificate request.

    .DESCRIPTION
    Completes a pending certificate request.  The pending request should be
    stored in either the user's or the machine's certificate store.

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
function Complete-CertificateRequest {
    [CmdletBinding()]
    param (
            [Parameter(Mandatory=$true)]
            [ValidateNotNullOrEmpty()]
            [System.IO.FileInfo]
            # The path to a certificate response file received from a
            # certification authority.
            $CACertificateResponse,

            [Parameter(Mandatory=$false)]
            [ValidateSet("Machine", "User")]
            [string]
            # The certificate store in which to look for the pending certificate request.
            $CertifcateStore
          )

    BEGIN {
        # The X509CertificateEnrollmentContext enum specifies that
        # "ContextMachine" is 0x2, which means store the certificate in the
        # Machine store.
        # http://msdn.microsoft.com/en-us/library/windows/desktop/aa379399.aspx
        $X509CertEnrollmentContextUser = 1
        $X509CertEnrollmentContextMachine = 2

        # The EncodingType enum value specifying that the encoding should
        # be represented in a specified format.
        #
        # http://msdn.microsoft.com/en-us/library/windows/desktop/aa374936.aspx
        $XCNCryptStringBase64Header = 0x0

        # InstallResponseRestrictionFlags =
        #   AllowNone                   = 0x00000000,
        #   AllowNoOutstandingRequest   = 0x00000001,
        #   AllowUntrustedCertificate   = 0x00000002,
        #   AllowUntrustedRoot          = 0x00000004
        #
        # http://msdn.microsoft.com/en-us/library/windows/desktop/aa376782.aspx
        $InstallResponseRestrictionFlags = 0x4
    }

    PROCESS {
        if ((Test-Path $CACertificateResponse) -eq $false) {
            Write-Error "$CACertificateResponse does not exist."
            return
        }
        Write-Verbose "Found file $CACertificateResponse"

        $response = Get-Content $CACertificateResponse

        $enrollment = New-Object -ComObject "X509Enrollment.CX509Enrollment.1"
        switch ($CertifcateStore) {
            "Machine" { $enrollment.Initialize($X509CertEnrollmentContextMachine) }
            "User"    { $enrollment.Initialize($X509CertEnrollmentContextUser)    }
        }

        # http://msdn.microsoft.com/en-us/library/windows/desktop/aa378051.aspx
        # The $null value is the optional password field, which you typically
        # don't want for a server certificate.
        $enrollment.InstallResponse($InstallResponseRestrictionFlags,
                                    $response,
                                    $XCNCryptStringBase64Header,
                                    $null)

        if ($enrollment.Status.ErrorText -match "successful") {
            Write-Host "Certificate Request completed successfully."
        } else {
            Write-Host $enrollment.Status.ErrorText
        }
    }
}
