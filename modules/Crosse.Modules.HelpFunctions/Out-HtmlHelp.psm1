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
            $OutputDir="$(Get-Location)\Help",

            [switch]
            # Whether to copy the template CSS file to the output directory.
            $CopyCssTemplate=$true
          )

    BEGIN {
        if ((Test-Path $OutputDir) -eq $false) {
            Write-Host "Making output directory"
            mkdir $OutputDir | Out-Null
            Write-Host "PSScriptRoot = $PSScriptRoot"
        }

        if ($CopyCssTemplate) {
            Write-Host "Copying template CSS file to output directory"
            copy "$PSScriptRoot\powershell-help.css" $OutputDir
        }
    }

    PROCESS {
        $fileName = Join-Path $OutputDir $($HelpInfo.Name + ".html")
        Write-Host "Writing help for `"$($HelpInfo.Name)`" to file $fileName"

        $doc = New-Object Crosse.Net.HtmlDocument
        $doc.Title = "$($HelpInfo.Name)"
        $doc.StyleSheet = "powershell-help.css"
        
        $doc.WriteHeading(1, $HelpInfo.Name)

        $doc.WriteHeading(2, "SYNOPSIS")
        $doc.WriteDiv("cmdSynopsis", $HelpInfo.Synopsis)
        $doc.WriteLine()

        $doc.WriteHeading(2, "SYNTAX")
        foreach ($item in $HelpInfo.syntax.syntaxItem) {
            $doc.BeginDiv("cmdSyntax")
            $doc.Write($item.name + " ")

            foreach ($p in $item.parameter) {
                if ($p.required -ne "true") {
                    $doc.Write("[")
                }

                $doc.Write("-$($p.name)")
                if ($p.parameterValue -ne $null) {
                    $doc.Write(" <$($p.parameterValue)>")
                }

                if ($p.required -ne "true") {
                    $doc.Write("]")
                }

                $doc.Write(" ")
            }
            $doc.EndDiv()
            $doc.WriteLine()
            $doc.WriteBreak()
            $doc.WriteLine()
        }

        $doc.WriteHeading(2, "DESCRIPTION")
        $doc.BeginDiv("cmdDescription")
        foreach ($d in $HelpInfo.description) {
            $doc.WriteLine()
            $doc.WriteLine($d.Tag + $d.Text)
            $doc.NewParagraph()
            $doc.WriteLine()
        }
        $doc.EndDiv()
        $doc.WriteLine()

        $doc.WriteHeading(2, "PARAMETERS")
        foreach ($p in $HelpInfo.parameters.parameter) {
            $doc.BeginDiv("cmdParameter")
            $doc.WriteLine()
            $doc.BeginDiv("cmdParameterName")
            $doc.Write("-$($p.name) ")

            if ($p.parameterValue -eq "SwitchParameter") {
                $doc.Write("[")
            }

            $doc.Write("<$($p.parameterValue)>")

            if ($p.parameterValue -eq "SwitchParameter") {
                $doc.Write("]")
            }
            $doc.EndDiv()
            $doc.WriteLine()

            $doc.BeginDiv("cmdParameterDesc")
            foreach ($d in $p.description) {
                $doc.WriteLine($d.Tag + $d.Text)
                $doc.WriteBreak()
            }
            $doc.WriteLine()
            $doc.EndDiv()
            $doc.WriteLine()

            $doc.BeginTable("cmdParameterAttr")
            
            $doc.BeginTableRow()
            $doc.BeginTableColumn("cmdParameterAttrName")
            $doc.Write("Required?")
            $doc.EndTableColumn()
            $doc.BeginTableColumn("cmdParameterAttrValue")
            $doc.Write($p.required)
            $doc.EndTableColumn()
            $doc.EndTableRow()

            $doc.BeginTableRow()
            $doc.BeginTableColumn("cmdParameterAttrName")
            $doc.Write("Position?")
            $doc.EndTableColumn()
            $doc.BeginTableColumn("cmdParameterAttrValue")
            $doc.Write($p.position)
            $doc.EndTableColumn()
            $doc.EndTableRow()

            $doc.BeginTableRow()
            $doc.BeginTableColumn("cmdParameterAttrName")
            $doc.Write("Accept pipeline input?")
            $doc.EndTableColumn()
            $doc.BeginTableColumn("cmdParameterAttrValue")
            $doc.Write($p.pipelineInput)
            $doc.EndTableColumn()
            $doc.EndTableRow()

            $doc.EndTable()
            $doc.EndDiv()
            $doc.WriteLine()            
        }


        $doc.WriteHeading(2, "INPUTS")
        $doc.BeginDiv("cmdInputs")
        $doc.WriteLine($HelpInfo.inputTypes.inputType.type.name)
        $doc.EndDiv()

        $doc.WriteHeading(2, "OUTPUTS")
        $doc.BeginDiv("cmdOutputs")
        $doc.WriteLine($HelpInfo.returnValues.returnValue.type.name)
        $doc.EndDiv()

        $doc.WriteHeading(2, "EXAMPLES")
        foreach ($e in $HelpInfo.examples.example) {
            $doc.BeginDiv("cmdExample")
            $doc.WriteLine()
            $doc.WriteLine($e.title)
            $doc.WriteBreak()
            $doc.WriteLine()

            foreach ($c in $e.code) {
                $doc.WriteLine($c)
                $doc.WriteBreak()
                $doc.WriteLine()
            }
            
            foreach ($r in $e.remarks) {
                if (![String]::IsNullOrEmpty($r.Text)) {
                    foreach ($l in $r.Text.Split("`n")) {
                        $doc.WriteLine($l)
                        $doc.WriteBreak()
                        $doc.WriteLine()
                    }
                }
                $doc.WriteLine()
            }
            
            $doc.EndDiv()
            $doc.WriteLine()
        }

        $doc.WriteHeading(2, "RELATED LINKS")
        $doc.BeginDiv("cmdRelatedLinks")
        foreach ($link in $HelpInfo.relatedLinks.navigationLink) {
            if ([String]::IsNullOrEmpty($link.linkText)) { 
                $linkText = $link.uri
            } else { 
                $linkText = $link.linkText
            }

            $doc.WriteLink("cmdRelatedLink", $link.uri, $linkText)
            $doc.NewParagraph()
        }
        $doc.EndDiv()

        $doc.WriteBreak();
        $doc.WriteBreak();
        
        $doc.ToString() | Out-File $fileName -Encoding ASCII
    }

    END {
        Write-Host "Finished."
    }
}
