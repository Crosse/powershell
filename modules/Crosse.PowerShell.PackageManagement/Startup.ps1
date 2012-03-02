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
    $code = @"
using System;
using System.IO;
using System.IO.Packaging;

namespace Crosse.PowerShell.PackageManagement {
    public class PackageFile : IDisposable {
        public string FileName { get; internal set; }
        public string Creator { get { return this.Package.PackageProperties.Creator; } }
        public string Title { get { return this.Package.PackageProperties.Title; } }
        public string Subject { get { return this.Package.PackageProperties.Subject; } }
        public string Category { get { return this.Package.PackageProperties.Category; } }
        public string Keywords { get { return this.Package.PackageProperties.Keywords; } }
        public string Description { get { return this.Package.PackageProperties.Description; } }
        public string Version { get { return this.Package.PackageProperties.Version; } }
        public string Revision { get { return this.Package.PackageProperties.Revision; } }
        public DateTime? Created { get { return this.Package.PackageProperties.Created; } }
        public DateTime? Modified { get { return this.Package.PackageProperties.Modified; } }
        public string LastModifiedBy { get { return this.Package.PackageProperties.LastModifiedBy; } }
        public string Identifier { get { return this.Package.PackageProperties.Identifier; } }
        public System.IO.Packaging.Package Package { get; internal set; }

        public PackageFile(string fileName, FileMode mode) {
            this.Package = System.IO.Packaging.Package.Open(fileName, mode);
            this.FileName = fileName;
        }

        public void Close() {
            if (this.Package != null) {
                this.Package.Close();
            }
            this.Package = null;
            this.FileName = null;
        }

        public void Dispose() {
            Close();
        }
    }
}
"@
    try {
        Write-Warning "Could not find $dll.  Compiling..."
        Add-Type -TypeDefinition $code -OutputAssembly $dll -ReferencedAssemblies $assembly.FullName
        Add-Type -Path $dll
    } catch {
        throw
    } finally {
        Remove-Variable assembly, dll, code
    }
}
