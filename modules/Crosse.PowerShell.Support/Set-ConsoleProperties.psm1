function Set-ConsoleProperties {
    param (
            [Parameter(Mandatory=$false)]
            [ValidateNotNullOrEmpty()]
            [string]
            $FontName,

            [Parameter(Mandatory=$false)]
            [ValidateRange(8, 32)]
            [int]
            $FontSize = 12,

            [Parameter(Mandatory=$false)]
            [ValidateRange(0, [Int]::MaxValue)]
            [int]
            $HistoryBufferSize = 250,

            [Parameter(Mandatory=$false)]
            [switch]
            $HistoryDuplication = $false,

            [Parameter(Mandatory=$false)]
            [ValidateRange(0, [Int]::MaxValue)]
            [int]
            $HistoryBufferSize = 250,

            [Parameter(Mandatory=$false)]
            [switch]
            $QuickEditMode = $true,
          )

    if ([String]::IsNullOrEmpty($FontName)) {
        $font = Get-Font -Name "Source Code Pro Medium"
        if ($font -eq $null) {
            $font = Get-Font -Name "Consolas"
        }
    }

    Set-ItemProperty -Path HKCU:\Console -Name FaceName -Value $FontName
    Set-ItemProperty -Path HKCU:\Console -Name HistoryNoDup -Value [int](!$HistoryDuplication)
    Set-ItemProperty -Path HKCU:\Console -Name HistoryBufferSize -Value $HistoryBufferSize
    Set-ItemProperty -Path HKCU:\Console -Name QuickEdit -Value $QuickEdit

    Set-ItemProperty -Path HKCU:\Console -Name FontSize -Value 0x120000
    Set-ItemProperty -Path HKCU:\Console -Name ScreenBufferSize -Value 0xbb8008c
    Set-ItemProperty -Path HKCU:\Console -Name ScreenBufferSize -Value 0x32008c
}

function Get-Font {
    param (
            [Parameter(Mandatory=$false,
                ParameterSetName="Filter")]
            [ScriptBlock]
            $Filter,

            [Parameter(Mandatory=$true,
                ParameterSetName="Name")]
            [string]
            $Name
          )

    $null = [System.Reflection.Assembly]::LoadWithPartialName("System.Drawing")
    $fonts = New-Object System.Drawing.Text.InstalledFontCollection

    if ($Filter -eq $null) {
        return $fonts
    } else {
        return ($fonts | Where-Object -FilterScript $Filter)
    }
}

