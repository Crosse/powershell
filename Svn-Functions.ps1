# detect source control management software
function global:findscm {
    $scm = ''
    :selectscm foreach ($_ in @('svn', 'hg')) {
        $dir = (Get-Location).ProviderPath.trimEnd("\")
        while ($dir.Length -gt 3) {
            if (Test-Path ([IO.Path]::combine(($dir), ".$_"))) {
                $scm = $_
                break selectscm
            }
            $dir = $dir -Replace '\\[^\\]+$', ''
        }
    }
    return $scm
}

# draw output
function global:drawlines($colors, $cmd) {
    $scm = findscm
    if (!$cmd -or !$scm) { return }
    foreach ($line in (&$scm ($cmd).split())) {
        $color = $colors[[string]$line[0]]
        if ($color) {
            write-host $line -Fore $color
        } else {
            write-host $line
        }
    }
}

# svn stat
function global:st {
    drawlines @{ "A"="Magenta"; "D"="Red"; "C"="Yellow"; "G"="Blue"; "M"="Cyan"; "U"="Green"; "?"="DarkGray"; "!"="DarkRed" } "stat $args"
}
Write-Host "`tAdded st (svn stat) to global functions." -Fore White

# svn update
function global:su {
    drawlines @{ "A"="Magenta"; "D"="Red"; "U"="Green"; "C"="Yellow"; "G"="Blue"; } "up $args"
}
Write-Host "`tAdded su (svn update) to global functions." -Fore White

# svn diff
function global:sd {
    drawlines @{ "@"="Magenta"; "-"="Red"; "+"="Green"; "="="DarkGray"; } "diff $args" 
}
Write-Host "`tAdded sd (svn diff) to global functions." -Fore White
