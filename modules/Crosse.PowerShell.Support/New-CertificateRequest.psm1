function New-CertificateRequest {
    [CmdletBinding()]
    param (
            [Parameter(Mandatory=$true)]
            [ValidateNotNullOrEmpty()]
            [string]
            $CommonName,

            [Parameter(Mandatory=$true)]
            [ValidateNotNullOrEmpty()]
            [string]
            $OrganizationalUnit,

            [Parameter(Mandatory=$true)]
            [ValidateNotNullOrEmpty()]
            [string]
            $Organization,

            [Parameter(Mandatory=$true)]
            [ValidateNotNullOrEmpty()]
            [string]
            $Locality,

            [Parameter(Mandatory=$true)]
            [ValidateNotNullOrEmpty()]
            [string]
            $State,

            [Parameter(Mandatory=$true)]
            [ValidateNotNullOrEmpty()]
            [string]
            $Country,

            [Parameter(Mandatory=$false)]
            $FriendlyName,

            [int]
            [ValidateSet(2048, 4096, 8192, 16384)]
            $KeyLength=2048,

            [Parameter(Mandatory=$false)]
            [string]
            $Description
          )

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
    $ExtensionKeyUsage = New-Object -ComObject "X509Enrollment.CX509ExtensionKeyUsage.1"
    $ExtensionKeyUsage.InitializeEncode(0xF0)
    $ExtensionKeyUsage.Critical = $true

    #   X509v3 Extended Key Usage: 
    #       TLS Web Server Authentication
    #
    # http://msdn.microsoft.com/en-us/library/windows/desktop/aa378132.aspx
    #   XCN_OID_PKIX_KP_SERVER_AUTH = 1.3.6.1.5.5.7.3.1
    #       "The certificate can be used for OCSP authentication."
    # 
    $EnhancedKeyUsageOid = New-Object -ComObject "X509Enrollment.CObjectId.1"
    $EnhancedKeyUsageOid.InitializeFromValue("1.3.6.1.5.5.7.3.1")
    $EnhancedKeyUsageOids = New-Object -ComObject "X509Enrollment.CObjectIds.1"
    $EnhancedKeyUsageOids.Add($EnhancedKeyUsageOid)
    $EnhancedKeyUsage = New-Object -ComObject "X509Enrollment.CX509ExtensionEnhancedKeyUsage.1"
    $EnhancedKeyUsage.InitializeEncode($EnhancedKeyUsageOids)

    # The name of the cryptographic provider.  Specifying this also sets the 
    # ProviderType.  In this case, the ProviderType is XCN_PROV_RSA_SCHANNEL.
    $ProviderName = "Microsoft RSA SChannel Cryptographic Provider"

    # Ohhh, it burns.  This is an SDDL string specifying that the 
    # NT AUTHORITY\SYSTEM and the BUILTIN\Administrators group have 
    # basically all rights to the certificate request, and that
    # NT AUTHORITY\NETWORK SERVICE has Read, List, and Create Child rights.
    # See the big scariness here:
    # http://msdn.microsoft.com/en-us/library/windows/desktop/aa772285(v=vs.85).aspx
    $SecurityDescriptor = "D:PAI(A;;0xd01f01ff;;;SY)(A;;0xd01f01ff;;;BA)(A;;0x80120089;;;NS)"

    # This is a X509KeySpec enum value that states that the key can be used for signing.
    $XCNAtKeyExchange = 1
    
    # The X509CertificateEnrollmentContext enum specifies that 
    # "ContextMachine" is 0x2, which means store the certificate in the
    # Machine store.
    $ContextMachine = 2

    # http://msdn.microsoft.com/en-us/library/windows/desktop/aa374936.aspx
    $XCNCryptStringBase64RequestHeader = 0x3

    $key = New-Object -ComObject "X509Enrollment.CX509PrivateKey.1"
    $key.ProviderName = $ProviderName
    $key.KeySpec = $XCNAtKeyExchange
    $key.Length = $KeyLength
    $key.SecurityDescriptor = $SecurityDescriptor
    # If MachineContext is true, then store this key in the Machine certificate
    # store.  If false, store the key in the user's personal store.
    $key.MachineContext = $true

    $key.Create()
    if (! $key.Opened ) {
        return
    }

    $subject = "CN={0},OU={1},O={2},L={3},S={4},C={5}" -f 
                $CommonName,
                $OrganizationalUnit,
                $Organization,
                $Locality,
                $State,
                $Country

    $distinguishedName = New-Object -ComObject "X509Enrollment.CX500DistinguishedName.1"
    $distinguishedName.Encode($subject)

    $certreq = New-Object -ComObject "X509Enrollment.CX509CertificateRequestPkcs10.1"
    $certreq.InitializeFromPrivateKey($ContextMachine, $key, $null)
    $certreq.Subject = $distinguishedName
    $certreq.X509Extensions.Add($ExtensionKeyUsage)
    $certreq.X509Extensions.Add($EnhancedKeyUsage)
    $certreq.SMimeCapabilities = $true

    $enrollment = New-Object -ComObject "X509Enrollment.CX509Enrollment.1"
    $enrollment.InitializeFromRequest($certreq)
    $csr = $enrollment.CreateRequest($XCNCryptStringBase64RequestHeader)

    if ($key.Opened) {
        $key.Close()
    }

    return $csr
}

