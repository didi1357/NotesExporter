#cs ----------------------------------------------------------------------------

 AutoIt Version: 3.3.14.2
 Copyright:      Licenced under the terms of GPL v3 or later version
 Author:         Dietmar Malli
 Version:        v1.0.2

 Script Function:
	Kill nonsense with nonsense! I <3 M$!
	This script is intended to say YES to the question wheter to open a .jnt
	file or not and wait for such questions in the background.

#ce ----------------------------------------------------------------------------

While 0 <> 1 ;forever
  While 0 <> 1 ;more forever
    Sleep(100) ;wait 100ms between tries..
    $handle = 0
    $english = WinExists("[REGEXPTITLE:(.*Windows Journal*)]", "contain security hazards")
    ; Yes Microsoft, ust add a small dash in the title to annoy people trying to automate that thing :D
    $german = WinExists("[REGEXPTITLE:(.*Windows-Journal*)]", "Sicherheitsrisiken enthalten")
    If $english <> 0 Then
      $handle = $english
    EndIf
    If $german <> 0 Then
      $handle = $german
    EndIf
	If $handle <> 0 Then
      ExitLoop
    EndIf
  WEnd
  While 0 <> 1
    WinActivate($handle)
    $succeeded = WinWaitActive($handle, "", 1)
	If $succeeded <> 0 Then
      ExitLoop
	EndIf
  WEnd
  Send("{j}")
  Sleep(100) ;wait 100ms.. programm won't start again within I guess
WEnd
