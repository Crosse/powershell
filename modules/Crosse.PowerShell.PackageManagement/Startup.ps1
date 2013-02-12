################################################################################
#
# Copyright (c) 2012 Seth Wright <wrightst@jmu.edu>
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

[CmdletBinding()]
param  ( )

if ([AppDomain]::CurrentDomain.GetAssemblies() -match "WindowsBase") {
    Write-Verbose "Found WindowsBase"
} else {
    $assembly = Get-ChildItem -Path 'C:\Program Files\Reference Assemblies\Microsoft\Framework' `
                                    -Filter "WindowsBase.dll" `
                                    -Recurse -ErrorAction SilentlyContinue
    if ($assembly -eq $null) {
        $assembly = Get-ChildItem -Path 'C:\Program Files (x86)\Reference Assemblies\Microsoft\Framework' `
                                    -Filter "WindowsBase.dll" `
                                    -Recurse -ErrorAction SilentlyContinue
    }
    if ($assembly -eq $null) {
        throw New-Object System.IO.FileNotFoundException "Cannot find WindowsBase.dll"
    }
    Add-Type -Path $assembly.FullName
}

$assembly = [AppDomain]::CurrentDomain.GetAssemblies() -match "WindowsBase"

$dll = Join-Path $PSScriptRoot "types.dll"
if (Test-Path $dll) {
    Write-Verbose "Found precompiled $dll"
    Add-Type -Path $dll
    Remove-Variable assembly, dll
} else {
    try {
        Write-Verbose "Compiling $dll"
        $code = [String]::Join("`n", (Get-Content (Join-Path $PSScriptRoot "Package.cs")))
        Add-Type -Language CSharp -TypeDefinition $code -ReferencedAssemblies $assembly.FullName
    } catch {
        throw
    } finally {
        Remove-Variable assembly, dll, code
    }
}
