function Out-TextToSpeech {
    [CmdletBinding()]
    param (
            [Parameter(Mandatory=$true,
                ValueFromPipeline=$true)]
            [string]
            $TextToSpeak,

            [switch]
            $SpeakAsync=$false
          )

    # The Module Manifest for this module includes the System.Speech
    # assembly in the list of assemblies to load when the module manifest
    # loads.  If you don't want to do this, or want to use this as a
    # standalone function, uncomment the next line.

    # [Reflection.Assembly]::LoadWithPartialName('System.Speech') | Out-Null
    $synth = New-Object System.Speech.Synthesis.Speechsynthesizer

    if ($SpeakAsync) {
        $synth.SpeakAsync($TextToSpeak)
    } else {
        $synth.Speak($TextToSpeak)
    }
}

Set-Alias say Out-TextToSpeech -Scope Global
