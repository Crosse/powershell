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

    [Reflection.Assembly]::LoadWithPartialName('System.Speech') | Out-Null
    $synth = New-Object System.Speech.Synthesis.Speechsynthesizer

    if ($SpeakAsync) {
        $synth.SpeakAsync($TextToSpeak)
    } else {
        $synth.Speak($TextToSpeak)
    }
}

New-Alias say Out-TextToSpeech -Scope Global
