function Invoke-TimedScriptBlock {
    param (
            [ScriptBlock]
            $ScriptBlock
          )

    $sw = New-Object System.Diagnostics.Stopwatch
    $sw.Start()
    $ScriptBlock.Invoke()
    $sw.Stop()
    if ($sw.ElapsedMilliseconds -le 1000) {
        Write-Host -ForegroundColor Yellow -BackgroundColor Black ("Took {0} milliseconds" -f $sw.ElapsedMilliseconds)
    } else {
        Write-Host -ForegroundColor Yellow -BackgroundColor Black ("Took {0} seconds" -f [Math]::Round($sw.Elapsed.TotalSeconds, 2))
    }
}

New-Alias -Force -Name time -Value Invoke-TimedScriptBlock
Export-ModuleMember -Function * -Alias *
