################################################################################
# 
# $URL$
# $Author$
# $Date$
# $Rev$
#
# DESCRIPTION:  
#
# Redistribution and use in source and binary forms, with or without           
# modification, are permitted provided that the following conditions are met:  
#
#  1. Redistributions of source code must retain the above copyright notice,   
#     this list of conditions and the following disclaimer.                    
#  2. Redistributions in binary form must reproduce the above copyright        
#     notice, this list of conditions and the following disclaimer in the      
#     documentation and/or other materials provided with the distribution.     
# 
# THIS SOFTWARE IS PROVIDED BY THE AUTHOR AND CONTRIBUTORS ``AS IS'' AND ANY   
# EXPRESS OR IMPLIED WARRANTIES, INCLUDING, BUT NOT LIMITED TO, THE IMPLIED    
# WARRANTIES OF MERCHANTABILITY AND FITNESS FOR A PARTICULAR PURPOSE ARE       
# DISCLAIMED. IN NO EVENT SHALL THE AUTHOR OR CONTRIBUTORS BE LIABLE FOR ANY   
# DIRECT, INDIRECT, INCIDENTAL, SPECIAL, EXEMPLARY, OR CONSEQUENTIAL DAMAGES   
# (INCLUDING, BUT NOT LIMITED TO, PROCUREMENT OF SUBSTITUTE GOODS OR SERVICES; 
# LOSS OF USE, DATA, OR PROFITS; OR BUSINESS INTERRUPTION) HOWEVER CAUSED AND  
# ON ANY THEORY OF LIABILITY, WHETHER IN CONTRACT, STRICT LIABILITY, OR TORT   
# (INCLUDING NEGLIGENCE OR OTHERWISE) ARISING IN ANY WAY OUT OF THE USE OF     
# THIS SOFTWARE, EVEN IF ADVISED OF THE POSSIBILITY OF SUCH DAMAGE.            
# 
################################################################################


param ([string]$Path='')

$wc = New-Object Net.WebClient
$i2e = "http://it-exmgmt.ad.jmu.edu/imap2exchange/addConversion?mailbox="

if ( !(Test-Path $Path) ) {
    Write-Error "Path does not exist"
    exit
}

# Open the file and start processing.
$lines = Get-Content $Path

# Either the file was empty, or something else happened that 
# prevented us from reading it.  Bail.
if ( !($lines) ) {
    Write-Error "Could not read file"
    exit
}

foreach ( $line in $lines ) {
    if ($line.Length -gt 0) {
        Write-Host -NoNewLine "Adding $line to the queue: "
        $wc.DownloadString($($i2e) + $line)
    }
}

