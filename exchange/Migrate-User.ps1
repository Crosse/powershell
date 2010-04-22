################################################################################
# 
# $URL$
# $Author$
# $Date$
# $Rev$
# 
# DESCRIPTION:  Migrates a user's mailbox from one Exchange organization to 
#               another.
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


param ( [switch]$Install=$false,
        [switch]$Verbose=$false,
        [string]$Identity=$null,
        [string]$SourceForestDomainController=$null,
        [string]$TargetForestDomainController=$null,
        [switch]$Cleanup=$false,
        $Credential=$null)

if (Test-Path function:Migrate-User) {
    Remove-Item function:Migrate-User
}

function global:Migrate-User( $inputObject=$Null,
        [switch]$Verbose=$false,
        [string]$Identity=$null,
        [string]$SourceForestDomainController=$null,
        [string]$TargetForestDomainController=$null,
        [switch]$Cleanup=$false,
        $Credential=(Get-Credential) ) {
    BEGIN {
        # This has something to do with pipelining.
        # Let's call it "magic voodoo" for now.
        if ($inputObject) {
            Write-Output $inputObject | &($MyInvocation.InvocationName) -Identity $Identity
            break
        }

        if ([String]::IsNullOrEmpty($SourceForestDomainController)) {
            Write-Error "Please specify a source domain controller (-SourceForestDomainController)"
            exit
        }

        if ([String]::IsNullOrEmpty($TargetForestDomainController)) {
            Write-Error "Please specify a target domain controller (-TargetForestDomainController)"
            exit
        }
    }
    PROCESS {
        if ([String]::IsNullOrEmpty($Identity)) {
            Write-Error "Please specify an Identity to migrate."
            exit
        }

        $targetmbx = Get-Mailbox -Identity $Identity -DomainController $TargetForestDomainController -ErrorAction SilentlyContinue
        # First lookup the user's mailbox in the current forest.
        if ($targetmbx -eq $null) {
            Write-Error "No mailbox exists for user $Identity in the current forest."
        }

        # Next, find the source mailbox and save it.
        $sourcembx = Get-Mailbox -DomainController $SourceForestDomainController -Credential $Credential -Identity $Identity
        Write-Host "Preparing to migrate $($sourcembx.DistinguishedName)"

        # Set the target mailbox's EmailAddresses property to include the PrimarySMTP 
        # address of the source mailbox.
        $emailAddresses = $targetmbx.EmailAddresses
        Write-Host $emailAddresses
        if (!($emailAddresses.Contains("smtp:$($sourcembx.PrimarySmtpAddress.ToString())")) ) {
            $emailAddresses.Add("smtp:$($sourcembx.PrimarySmtpAddress.ToString())")
        }

        # Add the source mailbox's legacyExchangeDN as an X.500 
        # address to the target mailbox's address collection.
        $sourceLegacyExchangeDN = $sourcembx.LegacyExchangeDN

        if (![String]::IsNullOrEmpty($sourceLegacyExchangeDN) -and 
                !($emailAddresses.Contains("X500:$($sourceLegacyExchangeDN)")) {
            $emailAddresses.Add("X500:$($sourceLegacyExchangeDN)")
        }

        Set-Mailbox -Identity $Identity -EmailAddressPolicyEnabled:$False -EmailAddresses $emailAddresses `
            -DomainController $TargetForestDomainController

        Get-Mailbox -Identity $Identity -DomainController $TargetForestDomainController | Select EmailAddresses

        # Move the mailbox.
        Get-Mailbox -DomainController $SourceForestDomainController -Credential $Credential -Identity $Identity |  
            Move-Mailbox -TargetDatabase $targetmbx.Database -SourceForestGlobalCatalog $SourceForestDomainController `
                -SourceForestCredential $Credential -AllowMerge -IgnorePolicyMatch -DomainController $TargetForestDomainController

        # Reset the EmailAddresses to something approaching sanity.
        Set-Mailbox -Identity $Identity -EmailAddressPolicyEnabled:$True -EmailAddresses $emailAddresses `
            -ManagedFolderMailboxPolicy "Default Managed Folder Policy" -ManagedFolderMailboxPolicyAllowed `
            -DomainController $TargetForestDomainController

        Get-Mailbox -Identity $Identity -DomainController $TargetForestDomainController | Select EmailAddresses
    }
    END {
    }
}

if ($Install -eq $True) {
    Write-Host "Added Migrate-User to global functions." -Fore White
    exit
} else {
    Migrate-User -Identity:$Identity -Verbose:$Verbose `
        -SourceForestDomainController:$SourceForestDomainController `
        -TargetForestDomainController:$TargetForestDomainController -Credential:$Credential -Cleanup:$Cleanup
}
