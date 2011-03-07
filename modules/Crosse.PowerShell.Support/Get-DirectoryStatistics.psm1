function Get-DirectoryStatistics { 
    param (
            [Parameter(Mandatory=$true,
                ValueFromPipeline=$true)]
            [System.IO.DirectoryInfo]
            # Specifies a path to a location.
            $Path,

            [switch]
            # Whether to recurse through subdirectories.  The default is true.
            $Recurse=$true,

            [ValidateNotNullOrEmpty()] 
            [string]
            # Specifies a filter in the provider's format or language. The
            # value of this parameter qualifies the Path parameter. The syntax of the
            # filter, including the use of wildcards, depends on the provider. Filters are
            # more efficient than other parameters, because the provider applies them when
            # retrieving the objects, rather than having Windows PowerShell filter the
            # objects after they are retrieved.
            $Filter,

            [switch]
            # Allows the cmdlet to get items that cannot otherwise not be
            # accessed by the user, such as hidden or system files.
            $Force=$true,

            [ValidateNotNullOrEmpty()]
            [ValidateSet("B", "KB", "MB", "GB", "TB")]
            [string]

            # Specifies how the resulting number should be formatted.  Valid 
            # values are:
            # B
            # KB
            # MB
            # GB
            # TB
            $FormatAs="B",

            [Int32]
            # Specifies the number of places to round the number to.  The
            # default is 2.
            $RoundTo=2
          )

BEGIN { }
PROCESS {
    $rawsize = New-Object Int64
    $command = "Get-ChildItem -Path `$Path -Force:`$Force -Recurse:`$Recurse"
    if ([String]::IsNullOrEmpty($Filter) -eq $false) {
        $command += " -Filter `"$Filter`""
    }

    $result = New-Object PSObject -Property @{
        Path                = $Path
        Size                = $null
        SizeQualifier       = $FormatAs
        FormattedSize       = $null
        TotalDirectories    = 0
        TotalFiles          = 0
    }

    foreach ($item in (Invoke-Expression $command)) {
        if ($item -is [System.IO.FileInfo]) {
            $rawsize += $item.Length
            $result.TotalFiles++
        } elseif ($item -is [System.IO.DirectoryInfo]) {
            $result.TotalDirectories++
        }
    }

    switch ($FormatAs) {
        "KB" { $fsize = $rawsize/1KB }
        "MB" { $fsize = $rawsize/1MB }
        "GB" { $fsize = $rawsize/1GB }
        "TB" { $fsize = $rawsize/1TB }
    }

    $result.Size = $rawsize
    $result.FormattedSize = "{0:N$RoundTo}" -f $fsize + "$FormatAs"
    $result
}
END { }
}
