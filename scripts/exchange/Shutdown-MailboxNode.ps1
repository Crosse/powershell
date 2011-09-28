################################################################################
# 
# NAME  : Shutdown-Node.ps1
# AUTHOR: Seth Wright , James Madison University
# DATE  : 4/23/2009
# 
# DESCRIPTION : Shuts down a Clustered Mailbox Server.
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

$moveCmsCommand = ".\Move-CMS.ps1"

$Choice = Read-Host "Do you wish to (S)hutdown or (R)estart this node, or (C)ancel? [R | S | C]"

[int]$shutdown = $null

switch ($Choice.Substring(0,1).ToUpper()) {
    "C" {
        Write-Warning "User cancelled operation."
        Start-Sleep -Seconds 5
        exit 1    
    }
    "R" {
        # '4' is the value given to Win32Shutdown to reboot.
        # I hate magic numbers...why do I keep using them?
        $shutdown = 4
    }
    "S" {
        # '8' is the value given to Win32Shutdown to shutdown.
        $shutdown = 8
    }
    default {
        Write-Error "Invalid response."
        exit 1
    }
}

if ($shutdown) {
    if (!(& "$moveCmsCommand" -s)) {
        # Something went wrong.
        Write-Error "Move-CMS command returned an error.  Aborting shutdown..."
        exit 1
    }
    # Then, shutdown this node.
    Write-Warning "Shutting down or restarting this node..."
    Write-Warning "Press Control-C to cancel in the next 15 seconds."
    
    foreach ($i in 15..1) { 
        Write-Host -NoNewline "$($i)..."
        Start-Sleep -Seconds 1
    }
    
    $os = Get-WmiObject Win32_OperatingSystem
    $os.psbase.Scope.Options.EnablePrivileges = $true
    if ($shutdown -eq 4) {
        # User wants to reboot.
        $os.Reboot()
    } elseif ($shutdown -eq 8) {
        # User wants to shutdown.
        $os.Shutdown()
    }
} else {
    Write-Error "`$shutdown was not specified"
}