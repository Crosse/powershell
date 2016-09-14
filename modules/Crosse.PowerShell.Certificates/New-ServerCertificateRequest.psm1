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
    Creates a new server certificate request.

    .DESCRIPTION
    Creates a new server certificate request and returns the Base64-encoded
    request text suitable for submitting to a third-party certificate authority.
    The certificate request is also stored in the local machine's certificate store,
    and can be completed using the Complete-CertificateRequest cmdlet.

    .INPUTS
    None
        You cannot pipe objects to the cmdlet.

    .OUTPUTS
    System.String.  New-ServerCertificateRequest returns the generated X509,
    Base64-encoded certificate request that can be submitted to a third-
    party certificate authority.

    .EXAMPLE
    C:\PS> New-ServerCertificateRequest -CommonName "servername"
    -----BEGIN NEW CERTIFICATE REQUEST-----
    [...]
    -----END NEW CERTIFICATE REQUEST-----

    The above example generates a certificate request suitable for use as a
    server certificate.
#>
################################################################################
function New-ServerCertificateRequest {
    [CmdletBinding(DefaultParameterSetName="RSA")]
    param (
            [Parameter(Mandatory=$true)]
            [ValidateNotNullOrEmpty()]
            [string]
            # The common name of the entity. For a client certificate,
            # this could be the user's email address.
            $CommonName,

            [Parameter(Mandatory=$false)]
            [ValidateNotNullOrEmpty()]
            [string]
            # The department or organizational unit in charge of the entity.
            $OrganizationalUnit,

            [Parameter(Mandatory=$false)]
            [ValidateNotNullOrEmpty()]
            [string]
            # The organzation or company in charge of the entity.
            $Organization,

            [Parameter(Mandatory=$false)]
            [ValidateNotNullOrEmpty()]
            [string]
            # The city or locality in which the organization resides.
            $Locality,

            [Parameter(Mandatory=$false)]
            [ValidateNotNullOrEmpty()]
            [string]
            # The non-abbreviated state or province in which the organization
            # resides.
            $State,

            [Parameter(Mandatory=$false)]
            [ValidateNotNullOrEmpty()]
            [ValidateLength(2, 2)]
            [string]
            # The 2-letter abbreviation of the country in which the organization
            # resides.
            $Country,

            [Parameter(Mandatory=$false)]
            [ValidateNotNullOrEmpty()]
            [string[]]
            # An array of alternative DNS names that should be bound to the
            # certificate's public key.  This can only be used for a server
            # certificate currently.
            $SubjectAlternateNames,

            [Parameter(Mandatory=$false, ParameterSetName="RSA")]
            [switch]
            # Generate an RSA certificate request.
            $RSA,

            [Parameter(Mandatory=$false, ParameterSetName="ECC")]
            [switch]
            # Generate an ECC (ECDSA) certificate request.
            $ECC,

            [Parameter(Mandatory=$false)]
            [ValidateSet(256, 384, 521, 2048, 4096, 8192, 16384)]
            [int]
            # The length of the key in bits.
            # For RSA certificates, the default is 2048. Valid values for RSA certificates are 2048, 4096, 8192, and 16384.
            # For ECC certificates, the default is 256. Valid values for ECC certificates are 256, 384, or 521.
            $KeyLength = -1

          )

    if (!$PSBoundParameters.ContainsKey($PSCmdlet.ParameterSetName)) {
        $PSBoundParameters.Add($PSCmdlet.ParameterSetName, $true)
    }
    New-CertificateRequest -ServerCertificate @PSBoundParameters
}