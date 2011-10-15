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

################################################################################
<#
    .SYNOPSIS
    Encodes a string to its Base64 representation.

    .DESCRIPTION
    Encodes a string into its corresponding Base64-encoded value.

    .INPUTS
    System.String. The string to encode can be piped into this cmdlet.

    .OUTPUTS
    System.String.  ConvertTo-Base64 returns the base-64-encoded value
    of the input string.

    .EXAMPLE
    C:\PS> ConvertTo-Base64 "Test String"
    VGVzdCBTdHJpbmc=

    The above example encodes the text "Text String" to a Base64 representation.

    .EXAMPLE
    C:\PS> Get-ChildItem C:\Windows\notepad.exe | ConvertTo-Base64 | Out-File .\notepad_base64.txt

    The above example shows how to pipe files into ConvertTo-Bas64.  In order to
    ensure that the Base64 encoding is correct, pass in the System.IO.FileInfo
    representation of the file.
#>
################################################################################
function ConvertTo-Base64 {
    [CmdletBinding()]
    param
        (
         [Parameter(Mandatory=$true,
             ValueFromPipeline=$true)]
         # The string to convert to Base64.
         $InputObject
        );

    BEGIN {
        if ($InputObject -ne $null -and $InputObject.GetType() -eq [Object[]]) {
            $strings = New-Object System.Collections.ArrayList
        }
    }
    PROCESS {
        Write-Verbose "InputObject is a $($InputObject.GetType())"

        if ($InputObject.GetType() -eq [Object[]]) {
            foreach ($string in $InputObject) {
                $strings.Add($string) | Out-Null
            }
        } elseif ($InputObject.GetType() -eq [System.IO.FileInfo]) {
            $bytes = [System.IO.File]::ReadAllBytes($InputObject)
            [System.Convert]::ToBase64String($bytes, "InsertLineBreaks");
        }
    }
    END {
        if ($InputObject.GetType() -eq [Object[]]) {
            $joined = [String]::Join("`n", $strings.ToArray());
            $bytes  = [System.Text.Encoding]::UTF8.GetBytes($joined);
            [System.Convert]::ToBase64String($bytes, "InsertLineBreaks");
        }
    }
}

################################################################################
<#
    .SYNOPSIS
    Converts a string from Base64 to its unencoded value.

    .DESCRIPTION
    Converts a string from Base64 to its unencoded value.

    .INPUTS
    System.String.  You can pipe the Base64-encoded string into this cmdlet.

    .OUTPUTS
    System.String.  ConvertFrom-Base64 returns the original, unencoded
    value of a Base64-encoded string.

    .EXAMPLE
    C:\PS> ConvertFrom-Base64 "VGVzdCBTdHJpbmc="
    Test String

    The above example decodes the Bas64 text to its unencoded representation,
    "Test String".

    .EXAMPLE
    C:\PS> Get-Content .\notepad.txt | ConvertFrom-Base64 -OutputFile .\notepad.exe

    The above example shows how to pipe data into ConvertFrom-Base64 and write
    it back out to a file.  Because the unencoded data may contain non-printable
    characters, do not use Out-File.
#>
################################################################################
function ConvertFrom-Base64 {
        param (
            [Parameter(Mandatory=$true,
                ValueFromPipeline=$true)]
            [String[]]
            # The text to convert from Base64.
            $InputObject,

            [String]
            # If specified, write the unencoded bytes to this file.
            $OutputFile
          );

    BEGIN {
        if ($OutputFile -notlike "*\*") {
            $OutputFile = "$pwd" + "\" + "$OutputFile"
        } elseif ($OutputFile -like ".\*") {
            $OutputFile = $OutputFile -replace "^\.",$pwd.Path
        } elseif ($OutputFile -like "..\*") {
            $OutputFile = $OutputFile -replace "^\.\.",$(get-item $pwd).Parent.FullName
        } else {
            throw "Cannot resolve path!"
        }

        $strings = New-Object System.Collections.ArrayList
    }
    PROCESS {
        foreach ($string in $InputObject) {
            $strings.Add($string) | Out-Null
        }
    }
    END {
        $joined = [String]::Join("`n", $strings.ToArray())
        $bytes  = [System.Convert]::FromBase64String($joined)
        if ($OutputFile) {
            [System.IO.File]::WriteAllBytes($OutputFile, $bytes)
        } else {
            [System.Text.Encoding]::UTF8.GetString($bytes);
        }
    }
}
