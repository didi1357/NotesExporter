#cs ----------------------------------------------------------------------------

 AutoIt Version: 3.3.14.2
 Copyright:      Licenced under the terms of GPL v3 or later version
 Author:         Dietmar Malli

 Script Function:
	Kill nonsense with nonsense! I <3 M$!
	This script is intended to say YES to the question wheter to open a .jnt
	file or not and wait for such questions in the background.

#ce ----------------------------------------------------------------------------

While 0 <> 1 ;forever
  While 0 <> 1 ;more forever
    Sleep(100) ;wait 100ms between tries..
    $handle = WinExists("[REGEXPTITLE:(.*Windows Journal*)]", "contain security hazards")
	If $handle <> 0 Then
      ExitLoop
    EndIf
  WEnd
  While 0 <> 1
    WinActivate("[REGEXPTITLE:(.*Windows Journal*)]", "contain security hazards")
    $succeeded = WinWaitActive("[REGEXPTITLE:(.*Windows Journal*)]", "contain security hazards", 1)
	If $succeeded <> 0 Then
      ExitLoop
	EndIf
  WEnd
  Send("{j}")
  Sleep(100) ;wait 100ms.. programm won't start again within I guess
WEnd

