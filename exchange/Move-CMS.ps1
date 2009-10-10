################################################################################
# 
# $URL$
# $Author$
# $Date$
# $Rev$
#
# DESCRIPTION:  This script moves a CMS from one node to the other
#               in an automated fashion, and is suitable to use as a shutdown
#               script.  It was modeled after the 
#               MoveClusteredMailboxServerScript.ps1 script as seen here:
#               http://telnetport25.wordpress.com/2008/07/29/powershell-and-moving-ccr-mailbox-server-instances-powershell-script/
#
# 
# Copyright (c) 2009 Seth Wright
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

# This is used only in case the script file itself is called with parameters,
# so that parameter completion will work on the command line (and other reasons).
# If no parameters are passed to the script, it will just load the Move-CMS
# function into the current runspace.
param($CMSName="", [switch]$f, [switch]$s, [switch]$m)

# If the function already exists in this runspace, remove it so it 
# can be re-added below.
if (Test-Path function:Move-CMS) { 
    Remove-Item function:Move-CMS 
}

############################################################
# This function will be inserted into the current          #
# runspace.  It does the real work of this script.         #
############################################################
function global:Move-CMS($CMSName="", [switch]$f, [switch]$s, [switch]$m, $inputObject=$Null) {
    
    ########################################
    # This section executes only once      #
    # before the pipeline.                 #
    ########################################
    BEGIN {
    
        # This has something to do with pipelining.  
        # Let's call it "magic voodoo" for now.
        if ($inputObject) {
            Write-Output $inputObject | &($MyInvocation.InvocationName) -CMSName $CMSName
            break
        }

        # For those playing the home game, this does nothing but ensure that 
        # the Microsoft.Exchange.Management.PowerShell.Admin PSSnapin is loaded.
    
        $ExchangeSnapin = "MICROSOFT.EXCHANGE.MANAGEMENT.POWERSHELL.ADMIN"
        $emsLoaded = $False
        
        # Iterate through all the loaded snapins, searching for the Exchange snapin.
        foreach ($snapin in (Get-PSSnapin | Sort-Object -Property Name)) {
            if ($snapin.name.ToUpper() -eq $ExchangeSnapin) {
                # Done, we have the extension and it's loaded.
                $emsLoaded = $True
                break
            }
        }
        
        if (!($emsLoaded)) {
            # The Exchange snapin was not loaded, so see if the 
            # extension is at least registered with the system.
            foreach ($snapin in (Get-PSSnapin -registered | Sort-Object -Property Name)) {
                if ($snapin.name.ToUpper() -eq $ExchangeSnapin) {
                    # Found the snapin; add it to the environment.
                    trap { continue }
                    Add-PSSnapin $ExchangeSnapin
                    Write-Host "Exchange 2007 Powershell Extensions found and added to this session."
                    $emsLoaded = $True
                break
                }
            }
        }
        
        if (!($emsLoaded)) {
            # The Exchange snapin is not installed on this system.
            # Print an error and bail.
            Write-Error -Category NotInstalled `
                -RecommendedAction "Install Exchange 2007 Powershell Extensions" `
                -Message "Exchange 2007 Powershell Extensions are not installed.  Please install the Extensions and re-run this command."
            continue
        } else {
            Write-Host "Found Exchange 2007 Powershell Extensions.  Continuing..."
        }
        
        $Log = $null
        
        if (![System.Diagnostics.EventLog]::sourceExists("Move-CMS")) {
            $Log = [System.Diagnostics.EventLog]::CreateEventSource("Move-CMS", "Application")
        }
        $Log = New-Object System.Diagnostics.EventLog("Application", ".")
        $Log.Source = "Move-CMS"
        
        $logError   = [System.Diagnostics.EventLogEntryType]::Error
        $logWarn    = [System.Diagnostics.EventLogEntryType]::Warning
        $logInfo    = [System.Diagnostics.EventLogEntryType]::Information

    } # end BEGIN

    ########################################
    # This section executes for each       #
    # object in the pipeline.              #
    ########################################
    PROCESS {
        $CurrentNode = Get-Content Env:ComputerName
        
        # Get the cluster of which this node is a member.
        # First, check to see if a CMSName was given on the command line.
        if (!($CMSName)) {
            # CMSName was not provided on the command line, or it is 
            $CMSName = Get-MailboxServer | where { $_.RedundantMachines -eq $CurrentNode }
        } elseif ($CMSName.GetType() -eq [System.String]) {
            $CMSName = Get-MailboxServer -Identity $CMSName
        } 
        # The variable $CMSName should now be an instance of 
        # Microsoft.Exchange.Data.Directory.Management.MailboxServer
        
        if (!($CMSName)) {
            # Couldn't get the CMS name for some reason.  Bail.
            Write-Error "Node $CurrentNode is not a Clustered Mailbox Server Node" `
                -RecommendedAction "This script should be run on a CMS, or the -CMSName parameter specified." `
                -Category InvalidOperation
            exit 1
        }
        
        Write-Host "Clustered Mailbox Server Name:  $CMSName"
        
        # Find the active node of the CMS.
        # First, get a list of all nodes in the CMS.
        $OperationalMachines = (Get-ClusteredMailboxServerStatus -Identity $CMSName).OperationalMachines
        
        # $pattern is the regex pattern to use to look for node marked as 
        # <Active...> in the OperationalMachines array.
        $activePattern = "^(?<activenode>.*)\s+<Active.*"
        
        # Perform the regex match.  $match is a throw-away variable.
        $match = $OperationalMachines | where { $_ -match $activePattern }
        
        if (!($matches.activenode)) {
            # No regex matches were found.
            $message = "Cannot determine the Active Node of CMS $CMSName"
            $Log.writeEntry($message, $logError, 404)
            Write-Error $message 
            exit 1
        }

        # A regex match was found for the Active Node.
        $ActiveNode = $matches.activenode
        
        Write-Host "The Active Node of CMS $CMSName is $ActiveNode"
        
        # Get the passive node by examining the RedundantMachines array.
        $PassiveNode = $CMSName.RedundantMachines | % { if ($_ -notlike $ActiveNode) { $_ } }
        if (!($PassiveNode)) { 
            $message = "Cannot determine the Passive Node of CMS $CMSName"
            $Log.writeEntry($message, $logError, 404)
            Write-Error $message 
            exit 1
        }
        
        Write-Host "The Passive Node of CMS $CMSName is $PassiveNode"
        
        # Ping the target (passive) node to ensure that it's alive.
        # First, create a new Ping object.
        $ping = New-Object System.Net.NetworkInformation.Ping
        
        # Since System.Net.NetworkInformation.Ping.Send() will generate
        # a nasty exception if the ping fails, and since it's a .NET object
        # (and doesn't implement the "-ErrorAction" parameter), 
        # suppress printing exceptions for this call.
        $ErrorActionPreference = "SilentlyContinue"
        $reply = $null
        
        # Ping the computer to see if it is alive.
        [System.Net.NetworkInformation.PingReply]$reply = $ping.Send($PassiveNode)
        $ErrorActionPreference = "Continue"
        
        if ($reply.Status -eq [System.Net.NetworkInformation.IPStatus]::Success ) {
            Write-Host "Able to ping target node $PassiveNode."
        } else {
            $message = "Cannot ping target node $PassiveNode.  Will not attempt to move the CMS."
            $Log.writeEntry($message, $logError, 410)
            Write-Error $message 
            exit 1
        }
                
        # '-s' was specified, which means that the CMS shouldn't be moved
        # to the local node if it is currently the passive node.
        if ($s -and ($CurrentNode -ne $ActiveNode) ) {
            $message = "Local node $CurrentNode is already the passive node, and " + `
                        "'-s' was specified.  The CMS will not be moved to this node."
            $Log.writeEntry($message, $logWarn, 412)
            Write-Warning $message
            return $True
        }
    
        # Unless '-f' was specified, ask the user if they *really* want to do this.
        if (!($f)) {
            # Issue a warning to the user.
            $WarningMessage = "Proceeding with this script will render " + `
                                "your Exchange Server offline for a period of time. " + `
                                "This means that all Mailboxes stored on this " + `
                                "server will be unavailable."
            Write-Warning $WarningMessage
            
            $Choice = Read-Host "Move the CMS `"$($CMSName)`" from node $ActiveNode to $($PassiveNode)? [Y:Yes | N:No]"
                
            If ($Choice.ToLower() -eq "n") {
                # The user answered "no".  Bail.
                $message = "User cancelled the operation."
                $Log.writeEntry($message, $logWarn, 304)
                Write-Host $message 
                exit 2
            }
        }

        # Two things happened to get here.
        # First, the user either answered "yes", or '-f' was specified.
        # Second, '-s' was specified and this node *is* the active node, 
        # or '-s' wasn't specified at all.
        $message = "Moving CMSName from to $PassiveNode; this may take a while."
        $Log.writeEntry($message, $logInfo, 100)
        Write-Host $message 
        
        $error.clear()
        # Move the CMS from the active node to the passive.
        Move-ClusteredMailboxServer -Identity $CMSName `
                                  -MoveComment "Script Based Move" `
                                  -TargetMachine $PassiveNode `
                                  -Confirm:$False `
                                  -ErrorAction SilentlyContinue
        
        # Test the error status.  If an error occurred, print it and die.
        if ( !([String]::IsNullOrEmpty($error[0])) ) {
            $message = "Could not move the CMS:  $($error[0])"
            $Log.writeEntry($message, $logError, 403)
            Write-Error $message 
            exit 1
        }
        
        # Test replication by enumerating the Storage Groups' copy statuses.
        # There are some "magic numbers" in this section (that I don't like).
        # The "20" below will ensure a full 90+ seconds for the SGs to 
        # finish initializing (see details about replication and LLR).
        for ($i = 0; $i -le 20; $i++) {
            # Pre-seed the value of $healthy.  If one of the SGs is not healthy,
            # this value will get set to $false down below.
            $healthy = $True
            
            # Iterate through all of the Storage Groups on the CMS and get their Copy Status.
            foreach ( $sgCopyStatus in (Get-StorageGroupCopyStatus -Server $CMSName) ) {
                if ($sgCopyStatus.SummaryCopyStatus -ne "Healthy") {
                    # This particular SG has a status other than "Healthy".
                    $healthy = $false
                    Write-Warning "Storage Group `"$($sgCopyStatus.StorageGroupName)`" is in state $($sgCopyStatus.SummaryCopyStatus)"
                }
            }
            if ($healthy) {
                # Everything's okay; break out of the for-loop.
                continue
            } else {
                # Sleep for 5 seconds if any storage group was marked as not Healthy.
                Write-Warning "Sleeping for 5 seconds..."
                Start-Sleep -Seconds 5
            }
        }
        
        # Test $healthy again to see what the final verdict was.
        if (!($healthy)) {
            # One of the SGs still wasn't healthy after 90 seconds.  Alert the operator.
            $message = "Storage Group Copy Status is not Healthy.  Manual intervention may be required."
            $Log.writeEntry($message, $logError, 501)
            Write-Warning $message 
        } else {
            # Everything was marked as Healthy.  We're done.
            $message = "Move operation was successful."
            $Log.writeEntry($message, $logInfo, 200)
            Write-Host $message 
        }

        # Print out some information.
        Write-Host "Output of Get-ClusteredMailboxServerStatus:"
        Get-ClusteredMailboxServerStatus -Identity $CMSName
        Write-Host "Summary of Get-StorageGroupCopyStatus:"
        Get-StorageGroupCopyStatus -Server $CMSName | select Identity, SummaryCopyStatus, FailedMessage
        
    } # end PROCESS block

    ########################################
    # This section executes only once,     #
    # after the pipeline.                  #
    ########################################
    END {
        # Nothing to do!
    } # end END block
} # end Move-CMS() function
############################################################


# Last little bit:  If you call the script with actual parameters
# instead of just using it to load the Move-CMS function into your
# runspace, deal with that by calling the function with the specified parameters.
if ($f -or $s -or $m -or $CMSName) {
    Move-CMS $CMSName -s:$s -f:$f -m:$m
} else {
    Write-Host "Added Move-CMS to global functions." -Fore White
}
