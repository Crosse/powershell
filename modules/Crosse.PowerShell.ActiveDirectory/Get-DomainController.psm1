################################################################################
# 
# $Id$
#
# DESCRIPTION: This script will create a new dynamic contact object in
#              in Active Directory.  Please see 
#              http://www.ietf.org/rfc/rfc2589.txt for more information.
#              Note:  The target domain MUST be in a forest operating at the 
#              Windows 2003 forest functional level!
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


function Get-DomainController {
    return [System.DirectoryServices.ActiveDirectory.Domain]::GetCurrentDomain().FindDomainController()

    <#
        .SYNOPSIS
        Finds a domain controller in the current domain.

        .DESCRIPTION
        Finds a domain controller in the current domain.

        .INPUTS
        Get-DomainController does not accept any input.

        .OUTPUTS
        A System.DirectoryServices.ActiveDirectory.DomainController object.

        .EXAMPLE

        C:\PS> Get-DomainController


        Forest                     : contoso.com
        CurrentTime                : 9/1/2010 7:56:21 PM
        HighestCommittedUsn        : 115168107
        OSVersion                  : Windows Server 2003
        Roles                      : {}
        Domain                     : ad.contoso.com
        IPAddress                  : 134.126.13.68
        SiteName                   : Default-First-Site-Name
        SyncFromAllServersCallback :
        Name                       : dc.ad.contoso.com
        
        .LINK
        http://msdn.microsoft.com/en-us/library/system.directoryservices.activedirectory.domaincontroller.getdomaincontroller.aspx
#>
}
