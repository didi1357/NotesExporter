#cs ----------------------------------------------------------------------------

 AutoIt Version: 3.3.14.2
 Copyright:      Licenced under the terms of GPL v3 or later version
 Author:         Dietmar Malli
 Version:        v1.0.4

 Script Function:
	Kill nonsense with nonsense! I <3 M$!
	This script is intended to say YES to the question wheter to open a .jnt
	file or not and wait for such questions in the background.

  Using handles for language detection and activation did not work. Only god
  and some AutoIT devs may know why.
  So the code just exists two times... Really bad style, but hey...

#ce ----------------------------------------------------------------------------

While 0 <> 1 ;forever
  While 0 <> 1
    Sleep(100) ;wait 100ms..
    $succeeded = 0
    If WinExists("[REGEXPTITLE:(.*Journal*)]", "Sicherheitsrisiken enthalten") <> 0 Then
      WinActivate("[REGEXPTITLE:(.*Journal*)]", "Sicherheitsrisiken enthalten")
      $succeeded = WinWaitActive("[REGEXPTITLE:(.*Journal*)]", "Sicherheitsrisiken enthalten", 1)
      If $succeeded <> 0 Then
        ExitLoop
      EndIf
    EndIf

    If WinExists("[REGEXPTITLE:(.*Journal*)]", "contain security hazards") <> 0 Then
      WinActivate("[REGEXPTITLE:(.*Journal*)]", "contain security hazards")
      $succeeded = WinWaitActive("[REGEXPTITLE:(.*Journal*)]", "contain security hazards", 1)
      If $succeeded <> 0 Then
        ExitLoop
      EndIf
    EndIf
  WEnd

  Send("{j}")
WEnd
