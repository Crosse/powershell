function Get-LdapUser {
    [CmdletBinding(SupportsShouldProcess=$true,
            ConfirmImpact="High")]
    
    param (
            [Parameter(Mandatory=$true)]
            [string]
            $LdapPath,

            [Parameter(Mandatory=$false)]
            [string]
            $BindDN,

            [Parameter(Mandatory=$false)]
            [string]
            $BindPassword,

            [Parameter(Mandatory=$false)]
            [System.DirectoryServices.AuthenticationTypes]
            $AuthenticationType="Anonymous",

            [Parameter(Mandatory=$true,
                ValueFromPipeline=$true)]
            [string]
            $Filter,

            [Parameter(Mandatory=$false,
                ValueFromRemainingArguments=$true)]
            [String[]]
            $Attributes
          )
            
    BEGIN {
        if ($LdapPath.StartsWith("LDAP://") -eq $false) {
            throw "Incorrect LDAP Path syntax"
        }

        $error.Clear()
        $root = New-Object System.DirectoryServices.DirectoryEntry $domain, $userName, $password, $authenticationType
        if ($root -eq $null) {
            throw "Error connecting to LDAP server:  $($error[0])"
        }
    }
    PROCESS {
        $query = New-Object System.DirectoryServices.DirectorySearcher
        $query.SearchRoot = $root
        $query.Filter = $Filter
        if ($Attributes -ne $null) {
            $query.PropertiesToLoad.AddRange($Attributes)
        }
        $result = $query.FindAll()

        foreach ($user in $result) {
            $ldapUser = New-Object PSObject
            foreach ($attrName in $user.Properties.PropertyNames) {
                if ($user.Properties[$attrName].Count -eq 1) {
                    $attribute = $user.Properties[$attrName][0]
                } else {
                    $attribute = $user.Properties[$attrName]
                }
                $ldapUser = Add-Member -PassThru -InputObject $ldapUser -MemberType NoteProperty -Name $attrName -Value $attribute
            }
            $ldapUser
        }
    }
    END {
    }
}
