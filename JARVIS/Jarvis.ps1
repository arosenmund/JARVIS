#############JARVIS##################################################



Function Loading_Art($Seconds){
                                $I= 0
                                While($I -lt $Seconds){
                                Write-host     "#                #"
                                start-sleep -Milliseconds 50
                                Write-host     " #              # "
                                start-sleep -Milliseconds 50
                                Write-host     "  #            #  "
                                start-sleep -Milliseconds 50
                                Write-host     "   #          #   "
                                start-sleep -Milliseconds 50
                                Write-host     "    #        #    "
                                start-sleep -Milliseconds 50
                                Write-host     "     #      #     "
                                start-sleep -Milliseconds 50
                                Write-host     "      #    #      "
                                start-sleep -Milliseconds 50
                                Write-host     "       #  #       "
                                start-sleep -Milliseconds 50
                                Write-host     "        ##        "
                                start-sleep -Milliseconds 50
                                Write-host     "        ##        "
                                start-sleep -Milliseconds 50
                                Write-host     "       #  #       "
                                start-sleep -Milliseconds 50
                                Write-host     "      #    #      "
                                start-sleep -Milliseconds 50
                                Write-host     "     #      #     "
                                start-sleep -Milliseconds 50
                                Write-host     "    #        #    "
                                start-sleep -Milliseconds 50
                                Write-host     "   #          #   "
                                start-sleep -Milliseconds 50
                                Write-host     "  #            #  "
                                start-sleep -Milliseconds 50
                                Write-host     " #              # "
                                start-sleep -Milliseconds 50
                                Write-host     "#                #"
                                start-sleep -Milliseconds 50
                                                      $I++}
                                          
                                
                                 
                                 


                                 }


######SPEECH FUNCTION###Speech Engine - .net Microsoft Native########################################
Function Speak_It([string]$text){
                                    Add-Type -AssemblyName System.speech
                                    $speak = New-Object System.Speech.Synthesis.SpeechSynthesizer
                                    $speak.SelectVoice('Microsoft Zira Desktop')
                                    $speak.Speak($text)  
                            }
######################################################################################################
######GREETINGS FOR TIME OF DAY FUNCTION##############################################################
Function TOD_Greeting(){
                            $TOD = (get-date).TimeOfDay
                            $Hour = $TOD.Hours
                            
                            If($Hour -lt 12){
                                                    $greeting = "Good Morning Aaron! It's me! Girl Jarvis. I hope you had a great night."
                                                    
                                                    }
                            ElseIf($Hour -ge 12 -and $Hour -le 17){

                                                                    $greeting = "Good Afternoon Aaron! Girl Jarvis here. I am happy to see you."
                                                                   
                                                                   }
                            ElseIf($Hour -gt 17 -and $Hour -le 19){

                                                                      $greeting = "Good Evening Aaron! Girl Jarvis at your service. It's almost time to go home!"
                                                                              
                                                                              }
                            Else{ $greeting = "Hello Aaron!  You are here late and girl jarvis is getting sleepy. Don't work too hard!" } 

                            $greeting
                            Speak_It($greeting)
                            }


####OUTLOOK#######

#Get current unread messages  ....and maybe read them to me
function Start_Outlook(){
$TTS_Starting = "I am now starting outlook for you. It should just be a moment."
Speak_it($TTS_Starting)
Start-Process -FilePath "C:\Program Files (x86)\Microsoft Office\Office15\OUTLOOK.EXE" -WindowStyle Minimized
Loading_Art(15)
Add-Type -AssemblyName "Microsoft.Office.Interop.Outlook"
$Outlook = New-Object -ComObject Outlook.Application
$Namespace = $Outlook.GetNameSpace("MAPI")
$Inbox = $NameSpace.Folders.Item(2).Folders |where {($_.FolderPath) -eq "\\aaron.rosenmund@us.af.mil\Inbox"}
$UI_Num = $Inbox.UnReadItemCount 

$UI_string = "You have $UI_Num unread emails in your inbox. Check them at your leisure."


####Custome Searches#####Secuirty anti-virus won't let me read the body because of the com interface.
####From Tracie Oster####
####From Michael Esquivias###
####From Brandon Devault#####
 

Speak_It($UI_string)

}

#Get Today's Appointments


#####Open Chrome Stuff######################
function Chrome_Stuff(){
$web1 = "mail.google.com"
$UI_1 = "your google mail, "
$web2 = "www.sans.org"
$UI_2 = "new cyber security information from sans, "
$web3 = "calendar.google.com/calendar/r/month"
$UI_3 = "personal calendar,"
$web4 = "https://music.amazon.com/home?ie=UTF8&ref_=sv_dmusic_7"
$UI_4 = "or play some music"
$other_things = $UI_1 + $UI_2 + $UI_3 + $UI_4

start-process "C:\Program Files (x86)\Google\Chrome\Application\chrome.exe" -ArgumentList $web1,$web2,$web3,$web4 -WindowStyle Maximized


$UI_string = "I will go ahead and open up your chrome items for you. Please take a look at $other_things, as I prepare the rest of your desktop."

Speak_It($UI_string)
}

#####Open Internet Explorer Stuff###########




#####Open One Note##########################




Function First_Login(){

    Loading_Art(4)

    TOD_Greeting

    Chrome_Stuff

    Loading_Art(8)

    Start_Outlook

    $UI_String = "I am done for now.  But if there is anything else you need let me know.  Lady Jarvis, OUT!"
    Speak_It($UI_String)
}


