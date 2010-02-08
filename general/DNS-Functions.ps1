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
                    $null = $ip6Addresses.Add($ipAddress)
                } else {
                    Write-Host "$($entry.HostName) has IPv6 address $ipAddress"
                }
            } else {
                if ($q) {
                    $null = $ip4Addresses.Add($ipAddress)
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
