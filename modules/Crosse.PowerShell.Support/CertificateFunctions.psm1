function New-CertificateRequest {
    [CmdletBinding()]
    param (
            [Parameter(Mandatory=$false)]
            [ValidateSet(2048, 4096, 8192, 16384)]
            [int]
            # The length of the key.  The default is 2048.
            $KeyLength=2048,

            [Parameter(Mandatory=$true,
                ParameterSetName="ExplicitDN")]
            [ValidateNotNullOrEmpty()]
            [string]
            $SubjectName,

            [Parameter(Mandatory=$true,
                ParameterSetName="ImplicitDN")]
            [ValidateNotNullOrEmpty()]
            [string]
            # The "CN=" value of the certificates' Subject field.
            $CommonName,

            [Parameter(Mandatory=$false,
                ParameterSetName="ImplicitDN")]
            [ValidateNotNullOrEmpty()]
            [string]
            # The "OU=" value of the certificates' Subject field.
            $OrganizationalUnit,

            [Parameter(Mandatory=$false,
                    ParameterSetName="ImplicitDN")]
            [ValidateNotNullOrEmpty()]
            [string]
            # The "O=" value of the certificate's Subject field.
            $Organization,

            [Parameter(Mandatory=$false,
                    ParameterSetName="ImplicitDN")]
            [ValidateNotNullOrEmpty()]
            [string]
            # The "L=" value of the certificate's Subject field.
            $Locality,

            [Parameter(Mandatory=$false,
                    ParameterSetName="ImplicitDN")]
            [ValidateNotNullOrEmpty()]
            [string]
            # The "S=" value of the certificate's Subject field.
            $State,

            [Parameter(Mandatory=$false,
                    ParameterSetName="ImplicitDN")]
            [ValidateNotNullOrEmpty()]
            [ValidateLength(2, 2)]
            [string]
            # The "C=" value of the certificate's Subject field.
            $Country,

            [Parameter(Mandatory=$false)]
            [ValidateNotNullOrEmpty()]
            [string[]]
            # An array of alternative names that should be bound to the
            # certificate's public key.
            $SubjectAlternateNames,

            [Parameter(Mandatory=$false)]
            [string]
            # An optional friendly name for the certificate.
            $FriendlyName,

            [Parameter(Mandatory=$false)]
            [string]
            # An option description of the certificate.
            $Description,

            [Parameter(Mandatory=$false)]
            [ValidateSet("Machine", "User")]
            [string]
            $Context = "Machine"

#            [Parameter(Mandatory=$false)]
#            [ValidateSet("Server", "User", "SMIME")]
#            [string]
#            $Type = "Server"
          )

    BEGIN {
        # The "magic numbers" section.
        # There is no really good reason for all these to be in the BEGIN
        # block, except that it sets these contant-value-type things apart
        # from the actual code.

        # The name of the cryptographic provider.  Specifying this also sets the
        # key's ProviderType.  In this case, the ProviderType is
        # XCN_PROV_RSA_SCHANNEL.
        # http://msdn.microsoft.com/en-us/library/windows/desktop/aa379427.aspx
        $ProviderName = "Microsoft RSA SChannel Cryptographic Provider"

        # This is an SDDL string specifying that the
        # NT AUTHORITY\SYSTEM and the BUILTIN\Administrators group have
        # basically all rights to the certificate request, and that
        # NT AUTHORITY\NETWORK SERVICE has Read, List, and Create Child rights.
        # See the big scariness here:
        # http://msdn.microsoft.com/en-us/library/windows/desktop/aa772285.aspx
        $ServerCertificateSDDL = "D:PAI(A;;0xd01f01ff;;;SY)(A;;0xd01f01ff;;;BA)(A;;0x80120089;;;NS)"

        # This is a X509KeySpec enum value that states that the key can be
        # used for signing.
        # http://msdn.microsoft.com/en-us/library/windows/desktop/aa379409.aspx
        $X509KeySpecKeyExchange = 1

        # X509CertificateEnrollmentContext
        #   ContextUser                        = 0x1,
        #   ContextMachine                     = 0x2,
        #   ContextAdministratorForceMachine   = 0x3
        #
        # http://msdn.microsoft.com/en-us/library/windows/desktop/aa379399.aspx
        $X509CertEnrollmentContextUser = 1
        $X509CertEnrollmentContextMachine = 2

        # Requested Extensions:
        #   X509v3 Key Usage: critical
        #       Digital Signature, Non Repudiation, Key Encipherment, Data Encipherment
        #
        # http://msdn.microsoft.com/en-us/library/windows/desktop/aa379410.aspx
        # X509KeyUsageFlags:
        #   XCN_CERT_DIGITAL_SIGNATURE_KEY_USAGE   = 0x80
        #   XCN_CERT_NON_REPUDIATION_KEY_USAGE     = 0x40
        #   XCN_CERT_KEY_ENCIPHERMENT_KEY_USAGE    = 0x20
        #   XCN_CERT_DATA_ENCIPHERMENT_KEY_USAGE   = 0x10
        #   == 0xF0 (240) when OR'ed together.
        $RequestedExtensions = 0xF0

        #   X509v3 Extended Key Usage:
        #       TLS Web Server Authentication
        #
        # http://msdn.microsoft.com/en-us/library/windows/desktop/aa378132.aspx
        #   XCN_OID_PKIX_KP_SERVER_AUTH = 1.3.6.1.5.5.7.3.1
        #       "The certificate can be used for OCSP authentication."
        #
        $RequestedEnhancedExtensions = "1.3.6.1.5.5.7.3.1"

        # The AlternativeNameType enum value specifying that an item in the
        # SubjectAlternativeNames list is a DNS name.
        # http://msdn.microsoft.com/en-us/library/windows/desktop/aa374830.aspx
        $XCNCertAltNameDnsName = 3

        # The EncodingType enum value specifying that the encoding should
        # be represented as a Certificate Request.  It puts the
        # -----BEGIN NEW CERTIFICATE REQUEST-----
        # -----END NEW CERTIFICATE REQUEST-----
        # text before and after the CSR.
        #
        # http://msdn.microsoft.com/en-us/library/windows/desktop/aa374936.aspx
        $XCNCryptStringBase64RequestHeader = 0x3
    }

    PROCESS {
        # Create the private key.
        # http://msdn.microsoft.com/en-us/library/windows/desktop/aa378921.aspx
        $key = New-Object -ComObject "X509Enrollment.CX509PrivateKey.1"
        $key.ProviderName = $ProviderName
        $key.KeySpec = $X509KeySpecKeyExchange
        $key.Length = $KeyLength
        if ($Context -eq "Machine") {
            $key.MachineContext = $true
            # Use an SDDL appropriate for a server certificate.
            $SecurityDescriptor = $ServerCertificateSDDL
        } else {
            $key.MachineContext = $false
            # Get the default SDDL appropriate for a user certificate.
            # http://msdn.microsoft.com/en-us/library/windows/desktop/aa376741.aspx
            $info = New-Object -ComObject X509Enrollment.CCspInformation.1
            $info.InitializeFromName($ProviderName)
            $SecurityDescriptor = $info.GetDefaultSecurityDescriptor($false)
        }
        $key.SecurityDescriptor = $SecurityDescriptor

        $key.Create()
        if (!$key.Opened) {
            Write-Error "Could not create and open a private key."
            return
        }
        Write-Verbose "Created private key."

        # Initialize the Certificate Request.
        # http://msdn.microsoft.com/en-us/library/windows/desktop/aa377505.aspx
        $certreq = New-Object -ComObject "X509Enrollment.CX509CertificateRequestPkcs10.1"
        if ($Context -eq "Machine") {
            $certreq.InitializeFromPrivateKey($X509CertEnrollmentContextMachine, $key, $null)
        } else {
            $certreq.InitializeFromPrivateKey($X509CertEnrollmentContextUser, $key, $null)
        }

        # Build the Subject attribute.
        if ($SubjectName -eq $null) {
            $subject = "CN={0}" -f $CommonName

            if ([String]::IsNullOrEmpty($OrganizationalUnit) -eq $false) {
                $subject += ",OU={0}" -f $OrganizationalUnit
            }

            if ([String]::IsNullOrEmpty($Organization) -eq $false) {
                $subject += ",O={0}" -f $Organization
            }

            if ([String]::IsNullOrEmpty($Locality) -eq $false) {
                $subject += ",L={0}" -f $Locality
            }

            if ([String]::IsNullOrEmpty($State) -eq $false) {
                $subject += ",S={0}" -f $State
            }

            if ([String]::IsNullOrEmpty($Country) -eq $false) {
                $subject += ",C={0}" -f $Country
            }
        } else {
            $subject = $SubjectName
        }

        Write-Verbose "Subject: $Subject"
        $distinguishedName = New-Object -ComObject "X509Enrollment.CX500DistinguishedName.1"
        $distinguishedName.Encode($subject)
        if ($distinguishedName.Name -eq $null) {
            Write-Error "Could not encode Subject value.  Ensure that it is of the proper format."
            return
        }
        $certreq.Subject = $distinguishedName

        $ExtensionKeyUsage = New-Object -ComObject "X509Enrollment.CX509ExtensionKeyUsage.1"
        $ExtensionKeyUsage.InitializeEncode($RequestedExtensions)
        $ExtensionKeyUsage.Critical = $true
        # Add the requested extensions to the certificate request.
        $certreq.X509Extensions.Add($ExtensionKeyUsage)

        $EnhancedKeyUsageOid = New-Object -ComObject "X509Enrollment.CObjectId.1"
        $EnhancedKeyUsageOid.InitializeFromValue($RequestedEnhancedExtensions)
        $EnhancedKeyUsageOids = New-Object -ComObject "X509Enrollment.CObjectIds.1"
        $EnhancedKeyUsageOids.Add($EnhancedKeyUsageOid)
        $EnhancedKeyUsage = New-Object -ComObject "X509Enrollment.CX509ExtensionEnhancedKeyUsage.1"
        $EnhancedKeyUsage.InitializeEncode($EnhancedKeyUsageOids)
        # Add the requested enhanced usage extensions to the certificate request.
        $certreq.X509Extensions.Add($EnhancedKeyUsage)

        # If the user specified that the certificate should include
        # alternative names, add them to the CSR.
        $alternativeNames = $null
        if ($SubjectAlternateNames.Count -gt 0) {
            for ($i = 0; $i -lt $SubjectAlternateNames.Count; $i++) {
                $name = $SubjectAlternateNames[$i]
                $altName = New-Object -ComObject "X509Enrollment.CAlternativeName.1"
                $altName.InitializeFromString($XCNCertAltNameDnsName, $name)
                if ($alternativeNames -eq $null) {
                    $alternativeNames = New-Object -ComObject "X509Enrollment.CAlternativeNames.1"
                }
                $alternativeNames.Add($altName)
                Write-Verbose "SubjectAlternativeName: DNS:$($name)"
            }
            if ($alternativeNames.Count -gt 0) {
                $ExtensionAlternativeNames = New-Object -ComObject "X509Enrollment.CX509ExtensionAlternativeNames.1"
                $ExtensionAlternativeNames.InitializeEncode($alternativeNames)
                # Add the requested Subject Alternative Names to the certificate request.
                $certreq.X509Extensions.Add($ExtensionAlternativeNames)
            }
        }

        # The CX509Enrollment object is what actually puts the CSR into the
        # certificate store and prints out the CSR for submission to a CA.
        # http://msdn.microsoft.com/en-us/library/windows/desktop/aa377809.aspx
        $enrollment = New-Object -ComObject "X509Enrollment.CX509Enrollment.1"
        $enrollment.InitializeFromRequest($certreq)
        if ($enrollment.Request -eq $null) {
            Write-Error $enrollment.Status.ErrorText
            return
        }

        if ([String]::IsNullOrEmpty($Description) -eq $false) {
            $enrollment.CertificateDescription = $Description
        }

        if ([String]::IsNullOrEmpty($FriendlyName) -eq $false) {
            $enrollment.CertificateFriendlyName = $FriendlyName
        }

        Write-Verbose "Creating Certificate Request"
        $csr = $enrollment.CreateRequest($XCNCryptStringBase64RequestHeader)
        if ($csr -eq $null) {
            Write-Error "Could not create the CSR: $($enrollment.Status.ErrorText)"
            return
        } else {
            Write-Verbose "Certificate Request created."
        }

        if ($key.Opened) {
            Write-Verbose "Closing private key"
            $key.Close()
        }

        return $csr
    }
}

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
            $Context
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
        if ($Context -eq "Machine") {
            $enrollment.Initialize($X509CertEnrollmentContextMachine)
        } else {
            $enrollment.Initialize($X509CertEnrollmentContextUser)
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
