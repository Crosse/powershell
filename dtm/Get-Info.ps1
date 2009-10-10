################################################################################
# 
# $URL$
# $Author$
# $Date$
# $Rev$
# 
# DESCRIPTION: Retrieves information from a remote computer in a variety of ways.
#
# USAGE:  ps_GetInfo [-v] -Computer <ComputerName> [-AttributesFile <AttributesFile.xml>]
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
# If no parameters are passed to the script, it will just load the ps_GetInfo
# function into the current runspace.
param($Computer="", $AttributesFile="Attributes.xml", [switch]$v, $inputObject=$Null)

BEGIN { # This section executes before the pipeline.

    # This has something to do with pipelining.  
    # Let's call it "magic voodoo" for now.
    if ($inputObject) {
        Write-Output $inputObject | &($MyInvocation.InvocationName) -AttributesFile $AttributesFile
        break
    }
    
    # Check to ensure that the Quest snapin has been registered.
    # Iterate through all the loaded snapins, searching for the Quest snapin.
    foreach ($snapin in (Get-PSSnapin | Sort-Object -Property Name)) {
        if ($snapin.name.ToUpper() -eq "QUEST.ACTIVEROLES.ADMANAGEMENT") {
        # Done, we have the extension and it's loaded.
        $questLoaded = $True
        break
        }
    }
    if (!($questLoaded)) {
        # The Quest snapin was not loaded, so see if the 
        # extension is at least registered with the system.
        foreach ($snapin in (Get-PSSnapin -registered | Sort-Object -Property Name)) {
            if ($snapin.name.ToUpper() -eq "QUEST.ACTIVEROLES.ADMANAGEMENT") {
                # Found the snapin; add it to the environment.
                trap { continue }
                Add-PSSnapin Quest.ActiveRoles.ADManagement
                Write-Host "Quest Active Directory Management Extensions found and added to this session."
                $questLoaded = $True
                break
            }
        }
    }
    
    if (!($questLoaded)) {
        # The Quest snapin is not installed on this system.
        # Print an error and bail.
        Write-Error -Category NotInstalled `
        -RecommendedAction "Install Quest Active Directory Management Extensions" `
        -Message "Quest Active Directory Management Extensions are not installed.  Please install the Extensions and re-run this command."
        continue
    }

    # Create the $ping object just once for the entire pipeline, so it can be reused.
    $ping = New-Object System.Net.NetworkInformation.Ping
    
    # See if an Attributes file was specified.  If not, bail.
    if ( !($AttributesFile) -or !(Test-Path $AttributesFile) ) {
        Write-Host "No Attributes file specified, or path does not exist: $AttributesFile"
        # Commented out so that we fall through to the PROCESS block and print the usage 
        # information there.
        #continue
    } else {
        # If we got here, then the Attributes file exists.  Try to load it.
        [xml]$attrFile = Get-Content -Path $AttributesFile
        if (!($attrFile)) {
            # Oops, either the file specified wasn't an XML file, or there was
            # an error in the file somewhere.  Bail.
            Write-Host "Error in $AttributesFile"
            continue
        }
    }
    
    # If we made it this far, start chucking stuff down the pipeline into the PROCESS block.
} # end 'BEGIN'

PROCESS {   # This section executes for each object in the pipeline.
    
    # Did we get data, either from the pipeline or explicitly on the command line?
    # If not, print out some (arguably "useful") help.
    if ( (!($_) -and !($Computer)) -or (!($AttributesFile)) ) {
        @"
        
USAGE:  ps_GetInfo [-v] -Computer <Computer> [-AttributesFile <Attributes.xml>]
Retrieve information from a remote computer.

Example: ps_GetInfo -Computer "$Env:ComputerName" -AttributesFile .\Attributes.xml

-v                 Verbose.  Writes the current computer name to the console.
                   Useful to monitor the progress of a pipeline operation.

-Computer          The computer to which you want to connect.

-AttributesFile    An XML file containing the attributes to retrieve.
                   If not specified, look for a file called "Attributes.xml"
                   in the current directory.

"@

        return
    }
    
    # If we got data via the pipeline, assign it to a named variable 
    # to make things easier to read.
    if ($_) { $Computer = $_ }
    
    # Try to find the computer in AD.  I guess this isn't really needed,
    # but it's one more validation that the computer exists.
    if ($Computer.GetType().Name -eq "ArsComputerObject") {
        # A Computer Object was passed; set an interim variable to the ArsComputerObject.
        $objComputer = $Computer
    } elseif ($Computer.GetType().Name -eq "String") {
        # A bare string was passed.  Turn it into an ArsComputerObject and
        # find it in Active Directory, then return the DnsName.
        $objComputer = Get-QADComputer -Name $Computer
        if (!($objComputer)) {
            # The computer account wasn't found in AD.  Bail on this object.
            Write-Host "$Computer is not in Active Directory."
            return
        }
    } else {
        # We have no idea what this object is.  Bail.
        Write-Host "$Computer is not an ArsComputerObject or a String."
        return
    }
        
    # Construct a new generic object to represent the Computer and use 
    # Add-Member to add some generic properties to the object.
    $Computer = New-Object PSObject
    $Computer = Add-Member -PassThru -InputObject $Computer NoteProperty Name $objComputer.DnsName
    # This property specifies whether the computer was able to be pinged.
    Add-Member -InputObject $Computer NoteProperty Pingable $False
    # This property specifies whether the script could connect to WMI on the remote computer.
    Add-Member -InputObject $Computer NoteProperty ConnectViaWmi $False
    # This property records the ModifcationDate (WhenChanged) property from the AD object.
    # Useful to decide how stale the computer object is.
    Add-Member -InputObject $Computer NoteProperty ADModificationDate $objComputer.ModificationDate
    
    # Iterate over the WMI Classes and Properties specified in the Attributes file and 
    # add properties to the $Computer object corresponding to the various
    # attributes in the Attributes file.
    foreach ($class in $attrFile.Attributes.Wmi.Classes.Class) {
        foreach ($property in $class.Property) {
            # Make sure to Trim() off any random whitespace from the XML file.
            Add-Member -InputObject $Computer NoteProperty $property.Trim() $null
        }
    }
    
    # Iterate over the files specified in the Attributes file and
    # add properties to the $Computer object.
    foreach ($file in $attrFile.Attributes.Files.File) {
        # Trim() off any random whitespace from the XML file.
        $propName = $file.Name.Trim()
        Add-Member -InputObject $Computer NoteProperty $propName $null
    }
        
    # Output the current computer object if -v was specified.
    if ($v) { Write-Host $Computer.Name }
    
    # Since System.Net.NetworkInformation.Ping.Send() will generate
    # a nasty exception if the ping fails, and since it's a .NET object
    # (and doesn't implement the "-ErrorAction" parameter), 
    # suppress printing exceptions for this call.
    $ErrorActionPreference = "SilentlyContinue"
    $reply = $null
    # Ping the computer to see if it is alive.
    [System.Net.NetworkInformation.PingReply]$reply = $ping.Send($Computer.Name)
    $ErrorActionPreference = "Continue"
    
    # The following logic test uses the .NET Enumeration 
    # instead of a string match, just in case.
    if ($reply.Status -eq [System.Net.NetworkInformation.IPStatus]::Success ) {
        # The computer is pingable, so note that and do some more stuff.
        $Computer.Pingable = $True

        # Clear the $wmi variable.
        $wmi = $null
        
        # Iterate through all WMI classes and get the various attributes
        # specified in the Attributes.xml file.
        foreach ($class in $attrFile.Attributes.Wmi.Classes.Class) {
            $wmi = Get-WmiObject -ComputerName $Computer.Name -ErrorAction SilentlyContinue $class.Name
            
            if ($wmi) {
                # We were able to connect to WMI.  Note this and continue.
                $Computer.ConnectViaWmi = $True
                
                # This actually retrieves the requested properties 
                # from the specified WMI class and sets the relevant
                # property on the $Computer object.
                foreach ($property in $class.Property) {
                    # Get the value of the property, trimming off any whitespace
                    # that may have made it into the XML file.
                    $property = $property.Trim()
                    
                    # On the off-chance that this Property is of the type CimType.DateTime,
                    # convert it to a more friendly string.
                    if ($wmi.PSBase.Properties["$property"].Type -eq [System.Management.CimType]::DateTime) {
                        # Retrieve the property, convert it, and save it.
                        $Computer.$property = $wmi.ConvertToDateTime($wmi.$property)
                    } else {
                        # Retrieve the property and save it.
                        $Computer.$property = $wmi.$property
                    } # end if
                } # end foreach ($property)
            } # end if ($wmi)
        } # end foreach ($class)

        # Since getting file information is a WMI call, only attempt it 
        # if the initial connection to WMI was successful.
        if ($wmi) {
            # Iterate over all of the files specified in the Attributes file 
            # and record the relevant information.
            foreach ($f in $attrFile.Attributes.Files.File) {
                # Just like above, construct the basename for the property name
                $baseName = $f.Name.Trim()
                # The WMI call wants escaped backslashes.
                $file = $baseName.Replace("\", "\\")
                
                # Retrieve the file's metadata from WMI (hopefully).
                $fObj = Get-WmiObject -ComputerName $Computer.Name -Class CIM_Datafile -Filter "Name=`'$file`'" -ErrorAction SilentlyContinue
                
                # If the file was found...
                if ($fObj) {
                    # ...and it has a version number...
                    if ($fObj.Version) {
                        # ...write that to the $Computer object...
                        $Computer.$baseName = $fObj.Version
                    } else {
                        # ...or else just record the fact that it was found.
                        $Computer.$baseName = $True
                    }
                } else {
                    # The file wasn't found, so set the property to false.
                    $Computer.$baseName = $False
                } # end if
            } # end foreach ($f)
        } # end if ($wmi)
    } # end if (Pingable)
    
    # We're done with this computer, so output it to pass it on down the pipeline.
    Write-Output $Computer
}

END {   # This section executes only once, after the pipeline.
    # Of course, there's not much we need to do here.
} # end 'END'
