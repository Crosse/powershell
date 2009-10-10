################################################################################
# 
# $URL$
# $Author$
# $Date$
# $Rev$
#
# DESCRIPTION: Migrates a user's mailbox from one Exchange organization to 
# another.
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


param ( [switch]$Install=$false,
        [switch]$Verbose=$false,
        [string]$Identity=$null,
        [string]$SourceForestDomainController=$null,
        [string]$TargetForestDomainController=$null,
        [string]$TargetDatabase=$null,
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
        [string]$TargetDatabase=$null,
        [switch]$Cleanup=$false,
        $Credential=(Get-Credential) ) {
    BEGIN {
        # This has something to do with pipelining.
        # Let's call it "magic voodoo" for now.
        if ($inputObject) {
            Write-Output $inputObject | &($MyInvocation.InvocationName) -Identity $Identity
            break
        }

        if ([String]::IsNullOrEmpty($TargetDatabase)) {
            Write-Error "Please specify the target database (-TargetDatabase)"
            exit
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

        # First create the user's mailbox in the current forest.
        if ((Get-Mailbox -Identity $Identity -ErrorAction SilentlyContinue) -eq $null) {
            Write-Host "Creating Mailbox in target forest for $($Identity)."
            Enable-Mailbox -Database $TargetDatabase -Identity $Identity -DomainController $TargetForestDomainController
        }

        # Next, find the source mailbox and save it.
        $sourcembx = Get-Mailbox -DomainController $SourceForestDomainController -Credential $Credential -Identity $Identity
        Write-Host "Preparing to migrate $($sourcembx.DistinguishedName)"

        # Set the target mailbox's EmailAddresses property to include the PrimarySMTP 
        # address of the source mailbox.
        $emailAddresses = (Get-Mailbox -Identity $Identity -DomainController $TargetForestDomainController).EmailAddresses
        $emailAddresses
        if (!($emailAddresses.Contains("smtp:$($sourcembx.PrimarySmtpAddress.ToString())")) ) {
            $emailAddresses.Add("smtp:$($sourcembx.PrimarySmtpAddress.ToString())")
        }
        Set-Mailbox -Identity $Identity -EmailAddressPolicyEnabled:$False -EmailAddresses $emailAddresses `
            -DomainController $TargetForestDomainController

        Get-Mailbox -Identity $Identity -DomainController $TargetForestDomainController | Select EmailAddresses

        # Move the mailbox.
        Get-Mailbox -DomainController $SourceForestDomainController -Credential $Credential -Identity $Identity |  
            Move-Mailbox -TargetDatabase $TargetDatabase -SourceForestGlobalCatalog $SourceForestDomainController `
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
        -SourceForestDomainController:$SourceForestDomainController -TargetDatabase:$TargetDatabase `
        -TargetForestDomainController:$TargetForestDomainController -Credential:$Credential -Cleanup:$Cleanup
}
