################################################################################
# 
# $URL$
# $Author$
# $Date$
# $Rev$
# 
# DESCRIPTION:  Found these functions at http://poshcode.org/1186/.  These are
#               NOT MINE; poshcode.org states it is a "...repository of scripts
#               that are free for public use."
#
# 
# Copyright (c) 2009,2010 Seth Wright <wrightst@jmu.edu>
#
# Permission to use, copy, modify, and distribute this software for any
# purpose with or without fee is hereby granted, provided that the above
# copyright notice and this permission notice appear in all copies.
#
# THE SOFTWARE IS PROVIDED "AS IS" AND THE AUTHOR DISCLAIMS ALL WARRANTIES
# WITH REGARD TO THIS SOFTWARE INCLUDING ALL IMPLIED WARRANTIES OF
# MERCHANTABILITY AND FITNESS. IN NO EVENT SHALL THE AUTHOR BE LIABLE FOR
# ANY SPECIAL, DIRECT, INDIRECT, OR CONSEQUENTIAL DAMAGES OR ANY DAMAGES
# WHATSOEVER RESULTING FROM LOSS OF USE, DATA OR PROFITS, WHETHER IN AN
# ACTION OF CONTRACT, NEGLIGENCE OR OTHER TORTIOUS ACTION, ARISING OUT OF
# OR IN CONNECTION WITH THE USE OR PERFORMANCE OF THIS SOFTWARE.
#
################################################################################


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
