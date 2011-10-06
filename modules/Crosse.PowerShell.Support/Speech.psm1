function Out-TextToSpeech {
    param (
            [Parameter(Mandatory=$true,
                ValueFromPipeline=$true)]
            [string]
            $Text,

            [switch]
            $Async=$false
          )

    # The Module Manifest for this module includes the System.Speech
    # assembly in the list of assemblies to load when the module manifest
    # loads.  If you don't want to do this, or want to use this as a
    # standalone function, uncomment the next line.

    # [Reflection.Assembly]::LoadWithPartialName('System.Speech') | Out-Null
    $synth = New-Object System.Speech.Synthesis.SpeechSynthesizer

    if ($Async) {
        $synth.SpeakAsync($Text)
    } else {
        $synth.Speak($Text)
    }
}

Set-Alias say Out-TextToSpeech -Scope Global
