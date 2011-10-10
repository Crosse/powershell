function Publish-Item {
    [CmdletBinding()]
    param (
            [Parameter(Mandatory=$true)]
            [ValidateNotNullOrEmpty()]
            [string]
            $FilePath,

            [Parameter(Mandatory=$true)]
            [ValidateNotNullOrEmpty()]
            [string]
            $Destination,

            [Parameter(Mandatory=$false,
                ParameterSetName="SignFiles")]
            [switch]
            $AddAuthenticodeSignature,

            [Parameter(Mandatory=$true,
                ParameterSetName="SignFiles")]
            [ValidateNotNull()]
            [System.Security.Cryptography.X509Certificates.X509Certificate]
            $Certificate
        )

    if ((Test-Path $FilePath) -eq $false) {
        Write-Error "Could not find $FilePath"
        return
    }

    if ((Test-Path $Destination) -eq $false) {
        New-Item -Path $Destination -Type Directory
    }

    if (Test-Path $FilePath -PathType Leaf) {
        $fPath = (Get-Item $FilePath).Directory
    } else {
        $fPath = Get-Item (Resolve-Path $FilePath)
    }

    $publishPath = Get-Item (Resolve-Path $Destination)

    foreach ($file in (Get-ChildItem $FilePath -Recurse)) {
        if ((Test-Path $file.FullName -PathType Container) -eq $true) {
            # Skip over directories
            continue
        }

        $path = $file.DirectoryName.Replace($fPath.Parent.FullName, $publishPath.FullName)
        if ((Test-Path $path) -eq $false) {
            New-Item -Path $path -Type Directory
        }

        $newFile = Join-Path -Path $path -ChildPath $file.Name
        if ((Test-Path $newFile) -eq $true) {
            $newFile = Get-Item $newFile
            if ($file.LastWriteTime -gt $newFile.LastWriteTime) {
                Copy-Item $file.FullName -Destination $newFile.FullName
            }
        } else {
            Copy-Item $file.FullName -Destination $newFile
            $newFile = Get-Item $newFile
        }

        if ($AddAuthenticodeSignature -and $newFile.Extension -match "`.ps.*1") {
            $sig = Get-AuthenticodeSignature $newFile
            $timestampServers = @(
                    "http://timestamp.comodoca.com/authenticode",
                    "http://timestamp.verisign.com/scripts/timstamp.dll",
                    "http://www.startssl.com/timestamp"
                    )

            $i = 0
            while ($sig.Status -ne 'Valid') {
                Write-Verbose "Authenticode signature on $($newFile.Name) is invalid, resigning"
                $sig = Set-AuthenticodeSignature -Certificate $Certificate `
                        -FilePath $newFile.FullName `
                        -TimestampServer $timestampServers[$i]
                $i++
                if ($i -gt $timestampServers.Count) {
                    Write-Verbose "Cannot contact any timestamp servers.  $($newFile.Name) is signed, but not timestamped."
                    break
                }
            }
        }
        $newFile
    }
}
