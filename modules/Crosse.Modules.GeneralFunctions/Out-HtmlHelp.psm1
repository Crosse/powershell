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
            <h1>$(TidyString $HelpInfo.Name)</h1>
            <h2>SYNOPSIS</h2>
                <div class="cmdSynopsis">$(TidyString $HelpInfo.Synopsis)</div>
            <h2>SYNTAX</h2>
"@ | Out-File $fileName

        foreach ($item in $HelpInfo.syntax.syntaxItem) {
            $params = @"
                <div class="cmdSyntax">
                    $(TidyString $item.name) 
"@
            foreach ($p in $item.parameter) {
                if ($p.required -ne "true") {
                    $params += "["
                }

                $params += "-$(TidyString $p.name)"
                if ($p.parameterValue -ne $null) {
                    $params+= " &lt;$(TidyString $p.parameterValue)&gt;"
                }

                if ($p.required -ne "true") {
                    $params += "]"
                }

                $params += " "
            }
            $params | Out-File $fileName -Append
@"
                </div>
                <br />
"@ | Out-File $fileName -Append
        }

@"
            <h2>DESCRIPTION</h2>
            <div class="cmdDescription">
"@ | Out-File $fileName -Append

        foreach ($d in $HelpInfo.description) {
            $desc = "                "
            $desc += $d.Tag + $d.Text
            $desc += "`n                <p />"
            $desc | Out-File $fileName -Append
        }

@"
            </div>
            <h2>PARAMETERS</h2>
"@ | Out-File $fileName -Append


        foreach ($p in $HelpInfo.parameters.parameter) {
@"
            <div class="cmdParameter">
                <div class="cmdParameterName">-$(TidyString $p.name) [&lt;$(TidyString $p.parameterValue)&gt;]</div>
                <div class="cmdParameterDesc">
"@ | Out-File $fileName -Append

            foreach ($d in $p.description) {
                $desc = "                "
                $desc += $d.Tag + $d.Text
                $desc += "`n                <br />"
                $desc | Out-File $fileName -Append
            }
@"
                </div>
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

@"
            <h2>INPUTS</h2>
                <div class="cmdInputs">$(TidyString $HelpInfo.inputTypes.inputType.type.name)</div>
            <h2>OUTPUTS</h2>
                <div class="cmdOutputs">$(TidyString $HelpInfo.returnValues.returnValue.type.name)</div>
            <h2>EXAMPLES</h2>
"@ | Out-File $fileName -Append

        foreach ($e in $HelpInfo.examples.example) {
            $example = @"
                <div class="cmdExample">
                    $(TidyString $e.title)<br />

"@
            foreach ($c in $e.code) {
                $example += @"
                    $(TidyString $c)<br />
"@
            }
            
            foreach ($r in $e.remarks) {
                $example += @"
                    $(TidyString $r.Text)<br />
"@
            }
            
            $example += @"
                </div>
"@
            $example | Out-File $fileName -Append
        }

@"
            <h2>RELATED LINKS</h2>
            <div class="cmdRelatedLinks">
"@ | Out-File $fileName -Append

        foreach ($link in $HelpInfo.relatedLinks.navigationLink) {
            if ([String]::IsNullOrEmpty($link.linkText)) { 
                $linkText = TidyString $link.uri
            } else { 
                $linkText = TidyString $link.linkText
            }
@"
                <a href="$($link.uri)" class="cmdRelatedLink">$linkText</a>
                <p />
"@ | Out-File $fileName -Append
        }

@"
           </div>
           <br />
           <br />
        </body>
    </html>
"@ | Out-File $fileName -Append

    Set-Encoding $fileName
    }

    END {
        Pop-Location
    }
}

function TidyString {
    param (
            [Parameter(ValueFromPipeline=$true)]
            # The string to convert.
            $Object
          )

        if ($Object -eq $null) {
            return $null
        }

        $retval = $Object.ToString()
        $retval = $retval.Trim()
        while ($retval[-1] -eq "`n") {
            $retval.Remove($retval.Length -1, 1)
        }
        $retval = $retval.Replace("&", "&amp;")
        $retval = $retval.Replace("<", "&lt;")
        $retval = $retval.Replace(">", "&gt;")
        $retval = $retval.Replace("`n", "<br />`n")

        return $retval
}

