################################################################################
# 
# $URL$
# $Author$
# $Date$
# $Rev$
# 
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


param ([switch]$Install=$False)

if (Test-Path function:template) {
    Remove-Item function:template
}

function global:template($inputObject=$Null) {
    BEGIN {
        # This has something to do with pipelining.  
        # Let's call it "magic voodoo" for now.
        if ($inputObject) {
            Write-Output $inputObject | &($MyInvocation.InvocationName) -CMSName $CMSName
            break
        }
    }
    PROCESS {
    }
    END {
    }
}

if ($Install -eq $True) {
    Write-Host "Added template to global functions." -Fore White
    exit
} else {
    template
}
