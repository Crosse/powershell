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

function ConvertTo-Base64 {
    [CmdletBinding()]
    param
        (
         [Parameter(ValueFromPipeline=$true)]
         # The string to convert.
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
            [Parameter(ValueFromPipeline=$true)]
            [String[]]
            # The string to convert.
            $InputObject,

            [String]
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
}
