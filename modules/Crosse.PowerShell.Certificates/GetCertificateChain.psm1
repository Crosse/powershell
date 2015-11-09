function Get-CertificateChain {
    [CmdletBinding()]
    param (
            [Parameter(Mandatory=$true,
                ValueFromPipeline=$true)]
            [ValidateNotNullOrEmpty()]
            [string]
            # The host to which to connect.  This can be either the fully-qualified
            # domain name or the IP address of the server.
            $Hostname,

            [Parameter(Mandatory=$false)]
            [ValidateRange(1, 65536)]
            [int]
            # The port to which to connect.  The default is port 443.
            $Port=443
          )

    PROCESS {
        $client = New-Object System.Net.Sockets.TcpClient $Hostname, $Port
        if (!$client.Connected) {
            Write-Error "Could not connect to ${Hostname}:${Port}"
            return
        }

        Write-Verbose "Connected to ${Hostname}:${Port}"


        $sslStream = New-Object System.Net.Security.SslStream `
                    $client.GetStream(), `
                    $false, `
                    {
                        param (
                                [object]
                                $sender,

                                [System.Security.Cryptography.X509Certificates.X509Certificate]
                                $certificate,

                                [System.Security.Cryptography.X509Certificates.X509Chain]
                                $chain,

                                [System.Net.Security.SslPolicyErrors]
                                $sslPolicyErrors
                                )

                        if ($sslPolicyErrors -eq [System.Net.Security.SslPolicyErrors]::None) {
                            Write-Verbose "Certificate and chain is valid."
                        } else {
                            Write-Warning "Certificate and/or chain is invalid:  $sslPolicyErrors"
                        }

                        $global:sslCertificateChain = $chain
                        return $true
                    }

        $sslStream.AuthenticateAsClient($Hostname)

        $result = $null
        if ($global:sslCertificateChain -ne $null) {
            $global:sslCertificateChain.Build($sslStream.RemoteCertificate) | Out-Null
            $result = $global:sslCertificateChain.ChainElements | % { $_.Certificate }
            Remove-Variable $global:sslCertificateChain
        }

        $sslStream.Close()
        $client.Close()

        return $result
    }
}
