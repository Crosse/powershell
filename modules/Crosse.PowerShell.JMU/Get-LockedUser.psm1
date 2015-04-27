function Get-LockedUser {
    [CmdletBinding()]
    param (
            [Parameter(Mandatory=$False, ValueFromPipeline=$true)]
            [String]
            # The user to search for.  If no user is specified, all lockout events will be returned.
            $User,

            [Parameter(Mandatory=$false)]
            [ValidateSet("Lockout", "Exchange")]
            [String]
            $EventType = "Lockout"
          )

    switch ($EventType) {
        "Lockout" {
            Search-LockoutEvents -User $User -ComputerName computer_name -EventLogName ForwardedEvents
        }
        "Exchange" {
            Search-LockoutEvents -User $User -FuzzySearch -ComputerName excas1_name, excas2_name
        }
    }
}

