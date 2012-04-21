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

$assembly = Get-ChildItem -Path 'C:\Program Files\Reference Assemblies\Microsoft\Framework' -Filter "WindowsBase.dll" -Recurse
if ($assembly -eq $null) {
    throw New-Object System.IO.FileNotFoundException "Cannot find WindowsBase.dll"
}
Add-Type -Path $assembly.FullName

$dll = Join-Path $PSScriptRoot "types.dll"
if (Test-Path $dll) {
    Add-Type -Path $dll
    Remove-Variable assembly, dll
} else {
    try {
        $code = [String]::Join("`n", (Get-Content (Join-Path $PSScriptRoot "Package.cs")))
        $start = Get-Date
        Add-Type -Language CSharpVersion3 -TypeDefinition $code -ReferencedAssemblies $assembly.FullName
        $end = Get-Date
        Write-Host "Package.cs compilation took $(($end - $start).Milliseconds)ms."
    } catch {
        throw
    } finally {
        Remove-Variable assembly, dll, code
    }
}
