param (
        [Parameter(Mandatory=$true)]
        [ValidateNotNullOrEmpty()]
        [string]
        $From,

        [Parameter(Mandatory=$true)]
        [ValidateNotNullOrEmpty()]
        [string]
        $To,

        [Parameter(Mandatory=$true)]
        [ValidateNotNullOrEmpty()]
        [string]
        $Subject,

        [Parameter(Mandatory=$true)]
        [ValidateNotNullOrEmpty()]
        [string]
        $BodyTemplate,

        [Parameter(Mandatory=$false)]
        [switch]
        $BodyAsHtml=$false,

        [Parameter(Mandatory=$false)]
        [ValidateScript({ if ([String]::IsNullOrEmpty($_) -and
                            [String]::IsNullOrEmpty($PSEmailServer)) {
                                Write-Host "No SMTP Server specified"
                                return $false
                            } else { return $true }})]
        [string]
        $SmtpServer,

        [Parameter(Mandatory=$false)]
        [switch]
        $UseSsl=$false,

        [Parameter(Mandatory=$true)]
        [System.Collections.Hashtable]
        $TemplateSubstitutions
    )

$Body = $BodyTemplate
foreach ($key in $TemplateSubstitutions.Keys) {
    $Body = $Body.Replace($key, $TemplateSubstitutions[$key])
}

Send-MailMessage -From $From `
                 -To $To `
                 -Subject $Subject `
                 -Body $Body `
                 -BodyAsHtml:$BodyAsHtml `
                 -SmtpServer $SmtpServer `
                 -UseSsl:$UseSsl 
