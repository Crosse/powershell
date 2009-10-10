################################################################################
# 
# $URL$
# $Author$
# $Date$
# $Rev$
# 
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

# If the function already exists in this runspace, remove it so it 
# can be re-added below.
if (Test-Path function:Get-HostEntry) { 
    Remove-Item function:Get-HostEntry
}

############################################################
# This function will be inserted into the current          #
# runspace.  It does the real work of this script.         #
############################################################
<#
.Synopsis
    Returns, in UNIX 'host'-style, the name or IP addresses of a host.
.Description
    Given either an IP address or a hostname, this function will attempt to 
    resolve the entry and will display the output in a style similar to 
    UNIX's 'host' command.
.Parameter HostEntry
    The name or IP address to lookup.
.Parameter q
    If specified, only print the returned IP address or host name.
.Example
    PS> Get-HostEntry www.google.com
    www.l.google.com has address 74.125.91.104
    www.l.google.com has address 74.125.91.147
    www.l.google.com has address 74.125.91.99
    www.l.google.com has address 74.125.91.103
.Example 
    PS> Get-HostEntry 74.125.91.103
    qy-in-f103.google.com has address 74.125.91.103


.ReturnValue
    [String]

.Notes
NAME:      Get-HostEntry
AUTHOR:    Seth Wright
LASTEDIT:  4/21/2009
#Requires -Version 2.0
#>
function global:Get-HostEntry(
        [Parameter(Position=0, Mandatory=$true, ValueFromPipeline=$true)]$HostEntry="", 
        [switch]$q, [switch]$4, [switch]$6, $inputObject=$Null) {
    ########################################
    # This section executes only once      #
    # before the pipeline.                 #
    ########################################
    BEGIN {
        # This has something to do with pipelining.  
        # Let's call it "magic voodoo" for now.
        if ($inputObject) {
            Write-Output $inputObject | &($MyInvocation.InvocationName) -HostEntry $HostEntry
            break
        }
    }
    PROCESS {
        if ($_) { $HostEntry = $_ }
        
        $error.clear()
        $ErrorActionPreference = "SilentlyContinue"
        $entry = [System.Net.Dns]::GetHostEntry($HostEntry)
        $ErrorActionPreference = "Continue"
        
        if (!($entry) -or !([String]::IsNullOrEmpty($error[0]))) { 
            Write-Host "Host $HostEntry not found"
            return
        }
        
        $error.clear()
        
        if ($q) {
            $obj = New-Object PSObject
            $obj = Add-Member -PassThru -InputObject $obj NoteProperty HostName $entry.HostName
            Add-Member -InputObject $obj NoteProperty IPv4Addresses $null
            Add-Member -InputObject $obj NoteProperty IPv6Addresses $null
            $ip4Addresses = New-Object System.Collections.ArrayList
            $ip6Addresses = New-Object System.Collections.ArrayList
        }

        foreach ($ipAddress in $entry.AddressList) {
            if ($ipAddress.AddressFamily -eq [System.Net.Sockets.AddressFamily]::InterNetworkV6) {
                if ($q) {
                    $ip6Addresses.Add($ipAddress)
                } else {
                    Write-Host "$($entry.HostName) has IPv6 address $ipAddress"
                }
            } else {
                if ($q) {
                    $ip4Addresses.Add($ipAddress)
                } else {
                    Write-Host "$($entry.HostName) has address $ipAddress"
                }
            }
        }

        if ($q) {
            $obj.IPv4Addresses = $ip4Addresses.ToArray()
            $obj.IPv6Addresses = $ip6Addresses.ToArray()
            $obj
        }
    }
}

Write-Host "`tAdded Get-HostEntry to global functions." -Fore White
