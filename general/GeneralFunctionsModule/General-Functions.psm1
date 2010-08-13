################################################################################
# 
# $URL$
# $Author$
# $Date$
# $Rev$
# 
# Copyright (c) 2009,2010 Seth Wright <wrightst@jmu.edu>
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
    Get-Content -Encoding byte $Path -ReadCount $Width -totalcount $Bytes | % {
        $byte = $_
            if (($byte -eq 0).count -ne $Width)
            {
                $hex = $byte | % {
                    " " + ("{0:x}" -f $_).PadLeft(2,"0")}
                $char = $byte | % {
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
        4d 5a 90 00 03 00 00 00 04 00 00 00 ff ff 00 MZ..........��.
        00 b8 00 00 00 00 00 00 00 40 00 00 00 00 00 ...............
        e8 00 00 00 0e 1f ba 0e 00 b4 09 cd 21 b8 01 �.....�....�...
        4c cd 21 54 68 69 73 20 70 72 6f 67 72 61 6d L�.This.program
        20 63 61 6e 6e 6f 74 20 62 65 20 72 75 6e 20 .cannot.be.run.
        69 6e 20 44 4f 53 20 6d 6f 64 65 2e 0d 0d 0a in.DOS.mode....
        24 00 00 00 00 00 00 00 e2 b1 38 08 a6 d0 56 ........�.8..�V
        5b a6 d0 56 5b a6 d0 56 5b af a8 d2 5b ef d0 ..�V..�V...�.��
    #>

}

function ConvertTo-Base64 {
    param
        (
         [Parameter(Mandatory=$true, ValueFromPipeline=$true)]
         [string]
         # The string to convert.
         $String
        );

    $bytes  = [System.Text.Encoding]::UTF8.GetBytes($string);
    $encoded = [System.Convert]::ToBase64String($bytes); 

    return $encoded;

    <#
        .SYNOPSIS
        Converts a string to its base-64-encoded value.

        .DESCRIPTION
        Converts a string to its base-64-encoded value.

        .INPUTS
        None.  You cannot pipe objects to ConvertTo-Base64.

        .OUTPUTS
        System.String.  ConvertTo-Base64 returns the base-64-encoded value
        of the input string.

        .EXAMPLE
        C:\PS> ConvertTo-Base64 "Test String"
        VGVzdCBTdHJpbmc=
    #>
}

function ConvertFrom-Base64 {
        param (
            [Parameter(Mandatory=$true, ValueFromPipeline=$true)]
            [string]
            # The string to convert.
            $String
          );

    $bytes  = [System.Convert]::FromBase64String($string);
    $decoded = [System.Text.Encoding]::UTF8.GetString($bytes); 

    return $decoded;

    <#
        .SYNOPSIS
        Converts a string from base-64 to its unencoded value.

        .DESCRIPTION
        Converts a string from base-64 to its unencoded value.

        .INPUTS
        None.  You cannot pipe objects to ConvertFrom-Base64.

        .OUTPUTS
        System.String.  ConvertFrom-Base64 returns the original, unencoded
        value of a base-64-encoded string.

        .EXAMPLE
        C:\PS> ConvertFrom-Base64 "VGVzdCBTdHJpbmc="
        Test String
    #>
};


function Get-RemoteFile { 
    param (
            [Parameter(Mandatory=$true, ValueFromPipeline=$true)]
            [string]
            # The URL to retrieve.
            $Url,

            [switch]
            # Whether to return save the file to disk or return
            # the data as a System.String. The default is False.
            $AsString=$false
          )

    # Attempt to download the file if a URL was specified.
    $wc = New-Object Net.WebClient
        if ($AsString) {
            $file = $wc.DownloadString($Url)
                Write-Output $file
        } else {
            $FileName = [System.IO.Path]::Combine((Get-Location).Path, $Url.Substring($Url.LastIndexOf("/") + 1))
                Write-Host "Saving to file $FileName"
                $wc.DownloadFile($Url, $FileName)
        }
    $wc.Dispose()
}
