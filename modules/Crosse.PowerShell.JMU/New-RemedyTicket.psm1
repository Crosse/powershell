function New-RemedyTicket {
    [CmdletBinding()]
    param (
            [Parameter(Mandatory=$true)]
            [ValidateNotNullOrEmpty()]
            [string]
            $RemedyServer,

            [Parameter(Mandatory=$true)]
            [ValidateNotNullOrEmpty()]
            [string]
            $Schema,

            [Parameter(Mandatory=$true)]
            [ValidateNotNullOrEmpty()]
            [string]
            $Key,

            [Parameter(Mandatory=$true)]
            [ValidateNotNullOrEmpty()]
            [string]
            $Submitter,

            [Parameter(Mandatory=$true)]
            [ValidateSet(
                    "Desktop Services",
                    "HelpDesk",
                    "Information Systems",
                    "Lab Admin",
                    "Lap Ops",
                    "Network Engineering",
                    "PC Services",
                    "Security Engineering",
                    "Systems",
                    "TSEC",
                    "Telecom",
                    "Training"
                    )]
            [string]
            $InitialWorkGroup,

            [Parameter(Mandatory=$true)]
            [ValidateNotNullOrEmpty()]
            [ValidateSet(
                    "Desktop Services",
                    "HelpDesk",
                    "Information Systems",
                    "Lab Admin",
                    "Lap Ops",
                    "Network Engineering",
                    "PC Services",
                    "Security Engineering",
                    "Systems",
                    "TSEC",
                    "Telecom",
                    "Training"
                    )]
            [string]
            $PresentWorkGroup,

            [Parameter(Mandatory=$true)]
            [ValidateNotNullOrEmpty()]
            [string]
            $CustomerUserName,

            [Parameter(Mandatory=$false)]
            [AllowNull()]
            [string]
            $CustomerFirstName,

            [Parameter(Mandatory=$false)]
            [AllowNull()]
            [string]
            $CustomerLastName,

            [Parameter(Mandatory=$true)]
            [ValidateNotNullOrEmpty()]
            [string]
            $AssignedTo,

            [Parameter(Mandatory=$true)]
            [ValidateNotNullOrEmpty()]
            [string]
            $Summary,

            [Parameter(Mandatory=$true)]
            [ValidateNotNullOrEmpty()]
            [string[]]
            $InitialCallText,

            [Parameter(Mandatory=$true)]
            [ValidateNotNullOrEmpty()]
            [string]
            $Category,

            [Parameter(Mandatory=$true)]
            [ValidateNotNullOrEmpty()]
            [string]
            $Item,

            [Parameter(Mandatory=$true)]
            [ValidateNotNullOrEmpty()]
            [string]
            $SubItem,

            [Parameter(Mandatory=$true)]
            [ValidateNotNullOrEmpty()]
            [string]
            $RemedyEmailAddress,

            [Parameter(Mandatory=$true)]
            [ValidateNotNullOrEmpty()]
            [string]
            $FromAddress,

            [Parameter(Mandatory=$false)]
            [ValidateNotNullOrEmpty()]
            $SmtpServer = $PSEmailServer
          )
    BEGIN {
        $fmtDate = [DateTime]::Now.ToString('M/dd/yyyy h:mm:ss t\M')
        $transfersLog = @(
                "Created by New-RemedyTicket.psm1 on $fmtDate",
                "Transferred to $PresentWorkGroup on $fmtDate",
                "Assigned to $AssignedTo on $fmtDate"
                ) -join "`n"
        $emailBody = @"
#AR-Message-Begin                    Do Not Delete This Line
Schema: $Schema
Server: $RemedyServer
Key: $Key
Action: Submit
Format: Short
Submitter !2!: $Submitter
Initial Work Group !562000298!: $InitialWorkGroup
Present Work Group !1000000003!: $PresentWorkGroup
JMU e-ID !562000121!: $CustomerUserName
First Name !540000100!: $CustomerFirstName
Last Name !540000000!: $CustomerLastName
Source !540000900!: Web
Assigned to !4!: $AssignedTo
Summary !8!: [`$`$$Summary`$`$]
Initial Call Text !540001100!: [`$`$$($InitialCallText.Split("`n") -join "^q+")`$`$]
Category !1000000010!: $Category
Item !1000000011!: $Item
SubItem !1000000012!: $SubItem
Group Permissions !112!: $PresentWorkGroup
Assigned To and Ownership Log !562000170!: [`$`$$transfersLog`$`$]
Create OL !562000487!: Yes
#AR-Message-End                      Do Not Delete This Line
"@
    }

    PROCESS {
        Write-Verbose $emailBody
        Send-MailMessage -From $FromAddress -To $RemedyEmailAddress -Cc $FromAddress -Body $emailBody -SmtpServer $SmtpServer -Subject "Incoming: $Summary"
    }

    END {
    }
}
