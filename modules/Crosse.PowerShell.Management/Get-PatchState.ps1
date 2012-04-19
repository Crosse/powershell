[CmdletBinding()]
param (
        [string]
        [Parameter(Mandatory=$false,
            ValueFromPipeline=$true)]
        $ComputerName,

        [System.IO.FileInfo]
        $PatchesFile
      )

BEGIN {
    $patches = Import-Csv $PatchesFile
    $ping = New-Object System.Net.NetworkInformation.Ping
}
PROCESS {
    if ($ping.Send($ComputerName, 1000).Status -ne "Success") {
        Write-Warning "Could not ping remote computer $ComputerName"
        return
    }

    $os = Get-WmiObject -ComputerName $ComputerName `
                        -Class Win32_OperatingSystem `

    foreach ($patch in $patches) {
        if ($os.Version.StartsWith($patch.OSVersion)) {
            $installedPatch = Get-WmiObject -ComputerName $ComputerName `
                    -Class Win32_QuickFixEngineering `
                    -Filter "HotFixID = `'$($patch.HotFixId)`'"

            if ($installedPatch -eq $null) {
                Write-Verbose "Did not find HotFixID $($patch.HotFixId) on computer $ComputerName"
                $IsHotFixInstalled = $false
                $Description = $null
                $HotFixID = $null
                $InstalledOn = $null
            } else {
                Write-Verbose "Found HotFixID $($patch.HotFixId) on computer $ComputerName"
                $IsHotFixInstalled = $true
                $Description = $installedPatch.Description
                $HotFixID = $installedPatch.HotFixID
                $InstalledOn = $installedPatch.InstalledOn
            }

            $result = New-Object PSObject 
            $result = Add-Member -PassThru -InputObject $result -MemberType NoteProperty -Name "ComputerName" -Value $ComputerName
            Add-Member -InputObject $result -MemberType NoteProperty -Name "OSVersion" -Value $os.Version
            Add-Member -InputObject $result -MemberType NoteProperty -Name "OSDescription" -Value $os.Name.Split("|")[0]
            Add-Member -InputObject $result -MemberType NoteProperty -Name "HotFixID" -Value $patch.HotFixID
            Add-Member -InputObject $result -MemberType NoteProperty -Name "IsHotFixInstalled" -Value $IsHotFixInstalled
            Add-Member -InputObject $result -MemberType NoteProperty -Name "PatchDescription" -Value $Description
            Add-Member -InputObject $result -MemberType NoteProperty -Name "InstalledOn" -Value $InstalledOn
            $result
        }
    }
}
