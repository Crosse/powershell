#function Out-TcpSocket {
    [CmdletBinding()]
    param (
            [Parameter(Mandatory=$false,
                ValueFromPipeline=$true)]
            [string]
            $InputObject,

            [Parameter(Mandatory=$true)]
            [ValidateNotNullOrEmpty()]
            [string]
            # The host to which to connect.  This can be either the fully-qualified
            # domain name or the IP address of the server.
            $Hostname,

            [Parameter(Mandatory=$true,
                ValueFromPipeline=$true)]
            [ValidateRange(1, 65536)]
            [int]
            # The port to which to connect.  The default is port 443.
            $Port,

            [Parameter(Mandatory=$false)]
            [ValidateSet("SSL", "StartTLS")]
            [string]
            $EncryptionMethod,

            [Parameter(Mandatory=$false)]
            [string]
            $LineSeparator="`r`n",

            [Parameter(Mandatory=$false)]
            [int]
            $TimeoutInMilliseconds=200
          )

    BEGIN {
        $client = New-Object System.Net.Sockets.TcpClient $Hostname, $Port
        if (!$client.Connected) {
            Write-Error "Could not connect to ${Hostname}:${Port}"
            return
        }

        Write-Verbose "Connected to ${Hostname}:${Port}"

        $stream = $null
        if ($EncryptionMethod -eq "SSL") {
            $stream = New-Object System.Net.Security.SslStream `
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

            $stream.AuthenticateAsClient($Hostname)
        } else {
            $stream = $client.GetStream()
        }

        $stream.ReadTimeout = $TimeoutInMilliseconds

        if ($stream.DataAvailable -eq $false) {
            $start = Get-Date
            while (($(Get-Date) - $start).TotalMilliseconds -le $TimeoutInMilliseconds) {
                if ($stream.DataAvailable -eq $true) {
                    break
                }
            }
        }

        $Encoding = New-Object System.Text.ASCIIEncoding
        if ($stream.DataAvailable -eq $true) {
            Write-Verbose "Reading data from socket"

            while ($stream.DataAvailable -eq $true) {
                $buff = New-Object Byte[] 1024
                $count = $stream.Read($buff, 0, 1024)
                $Encoding.GetString($buff, 0, $count)
            }
        }
        Write-Verbose "Stream setup finished."
    }

    PROCESS {
        #if ([String]::IsNullOrEmpty($InputObject) -eq $false) {
            Write-Verbose "Sending data to socket"
            Write-Output $InputObject
            $stream.Write(
                    [System.Text.Encoding]::ASCII.GetBytes($InputObject),
                    0,
                    $InputObject.Length)
            $stream.Write(
                    [System.Text.Encoding]::ASCII.GetBytes($LineSeparator),
                    0,
                    $LineSeparator.Length)
            $stream.Flush()
        #}

        if ($stream.DataAvailable -eq $false) {
            $start = Get-Date
            while (($(Get-Date) - $start).TotalMilliseconds -le $TimeoutInMilliseconds) {
                if ($stream.DataAvailable -eq $true) {
                    break
                }
            }
        }

        if ($stream.DataAvailable -eq $true) {
            Write-Verbose "Reading data from socket"

            while ($stream.DataAvailable -eq $true) {
                $buff = New-Object Byte[] 1024
                $count = $stream.Read($buff, 0, 1024)
                $Encoding.GetString($buff, 0, $count)
            }
        }
    }

    END {
        if ($stream.DataAvailable -eq $false) {
            $start = Get-Date
            while (($(Get-Date) - $start).TotalMilliseconds -le $TimeoutInMilliseconds) {
                if ($stream.DataAvailable -eq $true) {
                    break
                }
            }
        }

        if ($stream.DataAvailable -eq $true) {
            Write-Verbose "Reading data from socket"

            while ($stream.DataAvailable -eq $true) {
                $buff = New-Object Byte[] 1024
                $count = $stream.Read($buff, 0, 1024)
                $Encoding.GetString($buff, 0, $count)
            }
        }
        if ($UseSSL -eq $true) {
            $sslStream.Close()
        }

        $client.Close()
    }
#}
