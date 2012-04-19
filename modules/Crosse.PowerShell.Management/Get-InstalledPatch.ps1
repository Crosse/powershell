param (
        [string]
        $ComputerName,

        [string]
        $HotFixID
      )

if ([String]::IsNullOrEmpty($ComputerName)) {
    $ComputerName = "localhost"
}

Get-WmiObject   -ComputerName $ComputerName `
                -Class Win32_QuickFixEngineering `
                -Filter "HotFixID = `'$HotFixId`'"
