[void][System.Reflection.Assembly]::LoadWithPartialName("System.Speech");

##Setup the speaker, this allows the computer to talk
$speaker = [System.Speech.Synthesis.SpeechSynthesizer]::new();
$speaker.SelectVoice("Microsoft Zira Desktop");

##Setup the Speech Recognition Engine, this allows the computer to listen
$speechRecogEng = [System.Speech.Recognition.SpeechRecognitionEngine]::new();

##Setup the verbal commands hello and exit
$grammar = [System.Speech.Recognition.GrammarBuilder]::new();
$grammar2 = [System.Speech.Recognition.GrammarBuilder]::new();
$grammar.Append("Hey, Lady Jay.");
$grammar2.Append("Exit");
$speechRecogEng.LoadGrammar($grammar);
$speechRecogEng.LoadGrammar($grammar2);

$speechRecogEng.InitialSilenceTimeout = 15
$speechRecogEng.SetInputToDefaultAudioDevice();
$cmdBoolean = $false;

while (!$cmdBoolean) {
    $speechRecognize = $speechRecogEng.Recognize();
    $conf = $speechRecognize.Confidence;
    $myWords = $speechRecognize.text;
    if ($myWords -match "Hey, Lady Jay." -and [double]$conf -gt 0.85) {
        $speaker.Speak("Hello, Sir.  How are you?")}
    if ($myWords -match "exit" -and [double]$conf -gt 0.85) {
        $speaker.Speak("Goodbye")       
        $cmdBoolean = $true;
    }
}

