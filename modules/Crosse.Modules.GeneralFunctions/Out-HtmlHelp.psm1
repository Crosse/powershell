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

function Out-HtmlHelp {
    param (
            [Parameter(Mandatory=$true,
                ValueFromPipeline=$true)]
            # The Help info to process.
            $HelpInfo,

            [string]
            # The output directory.
            $OutputDir="$(Get-Location)\Help"
          )

    BEGIN {
        if ((Test-Path $OutputDir) -eq $false) {
            mkdir $OutputDir | Out-Null
        }
        Push-Location
        cd $OutputDir
    }

    PROCESS {
        $fileName = $HelpInfo.Name + ".html"
@"
        <html>
        <head>
            <title>$($HelpInfo.Name)</title>
            <link rel="stylesheet" type="text/css" href="powershell-help.css" />
        </head>

        <body>
            <h1>$($HelpInfo.Name)</h1>
            <h2>SYNOPSIS</h2>
                <div class="cmdSynopsis">$($HelpInfo.Synopsis)</div>
            <h2>SYNTAX</h2>
                <div class="cmdSyntax">$(ConvertTo-HtmlEntities $($HelpInfo.Syntax | Out-String -width 2000))</div>
            <h2>DESCRIPTION</h2>
"@ | Out-File $fileName

        foreach ($d in $HelpInfo.description) {
@"
                <div class="cmdDescription">$(ConvertTo-HtmlEntities $d.Text)</div>
"@ | Out-File $fileName -Append

@"
            <h2>PARAMETERS</h2>
"@ | Out-File $fileName -Append
        }


        foreach ($p in $HelpInfo.parameters.parameter) {
@"
            <div class="cmdParameter">
                <div class="cmdParameterName">-$($p.name) [&lt;$($p.parameterValue)&gt;]</div>
"@ | Out-File $fileName -Append

            foreach ($d in $p.description) {
@"
                <div class="cmdParameterDesc">$($d.Text)</div>
                <table class="cmdParameterAttr">
                    <tr>
                        <td class="cmdParameterAttrName">Required?</td>
                        <td class="cmdParameterAttrValue">$($p.required)</td>
                    </tr>
                    <tr>
                        <td class="cmdParameterAttrName">Position?</td>
                        <td class="cmdParameterAttrValue">$($p.position)</td>
                    </tr>
                    <tr>
                        <td class="cmdParameterAttrName">Accept pipeline input?</td>
                        <td class="cmdParameterAttrValue">$($p.pipelineInput)</td>
                    </tr>
                </table>
            </div>
"@ | Out-File $fileName -Append
            }
        }

@"
            <h2>INPUTS</h2>
                <div class="cmdInputs">$(ConvertTo-HtmlEntities $HelpInfo.inputTypes.inputType.type.name)</div>
            <h2>OUTPUTS</h2>
                <div class="cmdOutputs">$(ConvertTo-HtmlEntities $HelpInfo.returnValues.returnValue.type.name)</div>
            <h2>EXAMPLES</h2>
"@ | Out-File $fileName -Append

        foreach ($e in $HelpInfo.examples.example) {
@"
            <div class="cmdExample">$(ConvertTo-HtmlEntities ($e | Out-String -width 2000))</div>
            </div>
"@ | Out-File $fileName -Append
        }

@"
            <h2>RELATED LINKS</h2>
            <div class="cmdRelatedLinks">
"@ | Out-File $fileName -Append

        foreach ($link in $HelpInfo.relatedLinks.navigationLink) {
@"
                <a href="$($link.uri)" class="cmdRelatedLink">$(if ([String]::IsNullOrEmpty($link.linkText)) { $link.uri } else { $link.linkText })</a>
                <br/>
"@ | Out-File $fileName -Append
        }

@"
           </div>
           <br />
           <br />
        </body>
    </html>
"@ | Out-File $fileName -Append
    }

    END {
        Pop-Location
    }
}

function ConvertTo-HtmlEntities {
    param (
            [Parameter(Mandatory=$true,
                ValueFromPipeline=$true)]
            # The string to convert.
            $Object
          )

        return $Object.ToString().Replace("&", "&amp;").Replace("<", "&lt;").Replace(">", "&gt;").Replace("`n", "<p>")
}
