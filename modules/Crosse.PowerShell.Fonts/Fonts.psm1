function Get-Font {
    [CmdletBinding()]
    param (
            [Parameter(Mandatory=$false,
                ValueFromPipeline=$true)]
            [ValidateNotNullOrEmpty()]
            [string]
            $Name,

            [Parameter(Mandatory=$false)]
            [ValidateSet("Bold", "Italic", "Regular", "Strikeout", "Underline")]
            [System.Drawing.FontStyle]
            $Style
          )

    $WindowsFontsDir = Join-Path -Path (Get-Content Env:\WINDIR) -ChildPath "Fonts"
    if ((Test-Path $WindowsFontsDir) -eq $false) {
        throw "Unable to find the Windows Fonts directory"
    }

    $fonts = Get-ChildItem (Join-Path -Path $WindowsFontsDir -ChildPath '*.?tf')
    foreach ($fontFile in $fonts) {
        Write-Verbose $fontFile
        try {
            $gtf = New-Object System.Windows.Media.GlyphTypeface($fontFile.FullName)
            if ([String]::IsNullOrEmpty($Name) -or $gtf.FamilyNames.Values -contains $Name) {
                $gtf
            }
        } catch { }
    }
}

function Install-Font {
    [CmdletBinding()]
    param (
            [Parameter(Mandatory=$true,
                ValueFromPipeline=$true)]
            [string[]]
            $Path,

            [switch]
            $Force
          )

    BEGIN {
        $WindowsFontsDir = Join-Path -Path (Get-Content Env:\WINDIR) -ChildPath "Fonts"
        if ((Test-Path $WindowsFontsDir) -eq $false) {
            throw "Unable to find the Windows Fonts directory"
        }

        $objShell = New-Object -ComObject "Shell.Application"
    }

    PROCESS {
        # This is pretty much one of those situations that seems like cheating, but
        # is the only way to do something because the functionality isn't exposed
        # via .NET.

        $fontFile = [System.IO.FileInfo](Resolve-Path $Path).Path
        $fName = $fontFile.Name

        $installedFont = Join-Path -Path $WindowsFontsDir -ChildPath $fName
        if (Test-Path $installedFont) {
            # Get the SHA256 hashes of each font file to determine if the files
            # are the same. This is a pretty simple check; we could probably do
            # a lot more to verify font versions, etc., but for now a simple
            # "is it different" check will suffice.
            # If you really wanted to check if two fonts are alike,
            # System.Windows.Media.GlyphTypeface would probably be your friend.
            # (Remeber to add 'PresentationCore' to RequiredAssemblies, too.)
            $sourceHash = (Get-FileHash -Path $fontFile.FullName).Hash
            $installedHash = (Get-FileHash -Path $installedFont).Hash

            if ($sourceHash -eq $installedHash) {
                if ($Force) {
                    Write-Verbose "Reinstalling $fontFile"
                } else {
                    Write-Warning "$fontFile is already installed; use -Force to reinstall."
                    return
                }
            }
        }

        # Either the user is forcing installation or this font is not
        # installed, so install it.
        $objFolder = $objShell.NameSpace($fontFile.DirectoryName)
        $objFolderItem = $objFolder.ParseName($fName)
        $objFolderItem.InvokeVerb("Install")
    }
}
