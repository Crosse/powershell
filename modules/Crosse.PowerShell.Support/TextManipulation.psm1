################################################################################
#
# Copyright (c) 2011 Seth Wright <seth@crosse.org>
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

function Get-HexDump {
    param (
            [Parameter(Mandatory=$true, ValueFromPipeline=$true)]
            [string]
            # The path to the file to view in hex.
            $Path,

            [int]
            # The width of the hex dump. The default is 10.
            $Width=10,

            [int]
            # The number of bytes to decode. The default is -1.
            $Bytes=-1
          );

    $OFS=""
    Get-Content -Encoding byte $Path -ReadCount $Width -totalcount $Bytes | Foreach-Object {
        $byte = $_
            if (($byte -eq 0).count -ne $Width)
            {
                $hex = $byte | Foreach-Object {
                    " " + ("{0:x}" -f $_).PadLeft(2,"0")}
                $char = $byte | Foreach-Object {
                    if ([char]::IsLetterOrDigit($_))
                    { [char] $_ } else { "." }}
                "$hex $char"
            }
    }

    <#
        .SYNOPSIS
        Gets the contents of a file and displays it as a hex dump.

        .DESCRIPTION
        Gets the contents of a file and displays it as a hex dump.
        The width and length of the dump can be configured.

        .INPUTS
        A System.String describing the path to a file can be passed
        via the pipeline.

        .OUTPUTS
        System.String.  Get-HexDump returns a formatted hex dump of
        the file specified.

        .EXAMPLE
        C:\PS> Get-HexDump $Env:WINDIR\explorer.exe -Width 15 -Bytes 150
        4d 5a 90 00 03 00 00 00 04 00 00 00 ff ff 00 MZ..........??.
        00 b8 00 00 00 00 00 00 00 40 00 00 00 00 00 ...............
        e8 00 00 00 0e 1f ba 0e 00 b4 09 cd 21 b8 01 ?.....?....?...
        4c cd 21 54 68 69 73 20 70 72 6f 67 72 61 6d L?.This.program
        20 63 61 6e 6e 6f 74 20 62 65 20 72 75 6e 20 .cannot.be.run.
        69 6e 20 44 4f 53 20 6d 6f 64 65 2e 0d 0d 0a in.DOS.mode....
        24 00 00 00 00 00 00 00 e2 b1 38 08 a6 d0 56 ........?.8..?V
        5b a6 d0 56 5b a6 d0 56 5b af a8 d2 5b ef d0 ..?V..?V...?.??
    #>

}
Set-Alias hd Get-HexDump -Scope Global

function Set-Encoding {
    [CmdletBinding()]
        param (
                [Parameter(Position=0,Mandatory=$true,ValueFromPipeline=$true)]
                [ValidateScript({
                    if((resolve-path $_).Provider.Name -ne "FileSystem") {
                        throw "Specified Path is not in the FileSystem: '$_'"
                    }
                    return $true
                    })]
                [Alias("Fullname","Path")]
                [string]
                # The path to the file to be re-encoded.
                $FilePath,

                [Parameter(Mandatory=$false)]
                [ValidateSet(   "Unicode", "UTF7", "UTF8", "UTF32",
                                "ASCII", "BigEndianUnicode", "Default",
                                "OEM")]
                [string]
                # Specifies the type of character encoding used in the file.
                # ASCII is the default.
                $Encoding="ASCII"
             )

    PROCESS {
        (Get-Content $FilePath) | Out-File -FilePath $FilePath -Encoding $Encoding -Force
    }

    <#
        .SYNOPSIS
        Takes a Script file or any other text file into memory
        and Re-Encodes it in the format specified.

        .EXAMPLE
        ls *.ps1 | Set-Encoding -Encoding ASCII

        .DESCRIPTION
        Written to provide an easy method to perform easy batch
        encoding, calls on the command Out-File with the -Encoding
        parameter and the -Force switch. Primarily to fix UnknownError
        status received when trying to sign non-ascii format files with
        digital signatures. Don't use on your MP3's or other non-text
        files :)
#>
}

function Remove-ExtraWhitespace {
    [CmdletBinding()]
    param (
            [Parameter(Position=0,Mandatory=$true,ValueFromPipeline=$true)]
            [ValidateScript({
                if ((Resolve-Path $_).Provider.Name -ne "FileSystem") {
                    throw "Specified Path is not in the FileSystem: '$_'"
                }
                return $true})]
            [Alias("Fullname","Path")]
            [string]
            # The path to the file to be stripped of hanging whitespace.
            $FilePath
          )

    PROCESS {
        (Get-Content $FilePath | Foreach-Object { $_.TrimEnd() }) | Out-File -FilePath $FilePath -Force -Encoding ASCII
    }
}
