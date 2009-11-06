################################################################################
# 
# $Id$
# 
# DESCRIPTION:  This script sets the 'EnableOabGenOnThisNode' registry key on 
#               a CCR node member to the active node in the cluster.
#               Reference http://technet.microsoft.com/en-us/library/bb266910.aspx
#               for more information.
#
# Copyright (c) 2009 Seth Wright (wrightst@jmu.edu)
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
################################################################################

$CurrentNode = Get-Content Env:ComputerName
$Log = $null

if (![System.Diagnostics.EventLog]::sourceExists("EnableOabGen")) {
    $Log = [System.Diagnostics.EventLog]::CreateEventSource("EnableOabGen", "Application")
}
$Log = New-Object System.Diagnostics.EventLog("Application", ".")
$Log.Source = "EnableOabGen"

$error.Clear()
$CMSName = Get-MailboxServer | where { $_.RedundantMachines -eq $CurrentNode }

if (!($CMSName) -or !([String]::IsNullOrEmpty($error[0]))) {
    $Log.writeEntry("Could not determine the CMSName: $($error[0])", [System.Diagnostics.EventLogEntryType]::Error, 404)
    return
}

# Find the active node of the CMS.
# First, get a list of all nodes in the CMS.
$error.Clear()
$OperationalMachines = (Get-ClusteredMailboxServerStatus -Identity $CMSName).OperationalMachines

if (!($OperationalMachines) -or !([String]::IsNullOrEmpty($error[0]))) {
    $Log.writeEntry("Could not determine OperationalMachines: $($error[0])", [System.Diagnostics.EventLogEntryType]::Error, 404)
    return
}

# $pattern is the regex pattern to use to look for node marked as 
# <Active...> in the OperationalMachines array.
$activePattern = "^(?<activenode>.*)\s+<Active.*"

# Perform the regex match.  $match is a throw-away variable.
$match = $OperationalMachines | where { $_ -match $activePattern }

if (!($matches.activenode)) {
    # No regex matches were found.
    $Log.writeEntry("Cannot determine the Active Node of CMS $CMSName", [System.Diagnostics.EventLogEntryType]::Error, 404)
    return
}

# A regex match was found for the Active Node.
$ActiveNode = $matches.activenode

$baseKey = 'HKLM:\SYSTEM\CurrentControlSet\Services\MSExchangeSA\Parameters\' + $CMSName.Name
$result = (Get-ItemProperty -Path $baseKey).EnableOabGenOnThisNode
if ($result -notlike $ActiveNode) {
    Set-ItemProperty -Path $baseKey -Name "EnableOabGenOnThisNode" -Value "$ActiveNode"
    
    $result = $null
    $result = (Get-ItemProperty -Path $baseKey).EnableOabGenOnThisNode
    
    $Log.writeEntry("The registry value `"EnableOabGenOnThisNode`" on $CurrentNode has been set to $result", [System.Diagnostics.EventLogEntryType]::Warning, 200)
}
