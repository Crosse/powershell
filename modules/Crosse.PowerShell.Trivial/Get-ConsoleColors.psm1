function Get-ConsoleColors {
    $bgcolor = $Host.UI.RawUI.BackgroundColor
    $colors = [Enum]::GetValues([ConsoleColor])
    foreach ($color in $colors) {
        $val = ($colors.Count - $color.ToString().Length) / 2
        $lside = [Math]::Floor($val)
        $rside = [Math]::Ceiling($val)
        $swatch = "{0,$lside}{1}{0,$rside}" -f " ", $color
        Write-Host $swatch -BackgroundColor $color -ForegroundColor (($color.value__ + ($colors.Count/2)) % $colors.Count)
    }
}
