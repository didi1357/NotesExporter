#cs ----------------------------------------------------------------------------

 AutoIt Version: 3.3.14.2
 Copyright:      Licenced under the terms of GPL v3 or later version
 Author:         Dietmar Malli

 Script Function:
	This script takes a .one file as parameter and will make OneNote export it
	as a pdf to the second paramter specified.
	It assumes that the file is not existing. (Overwrite dialogue not handled.)
	The save as dialogue title is expected to be in german.
	Search for "Speichern unter" if you need to change that.
	All this is done in a very strange way because I ran into many problems.
	  +AutoIT is buggy as ****
	  +OneNote doesn't offer a command line api. (This would make the whole
	   file and my hours spent here useless >.<)
	  +Generally automating modern M$ GUIs like the one of OneNote with AutoIt
	   feels like putting **** on top of ****.
    I'm quite sure this script will still not work in every case.
	Even after about 20 hours of testing.

#ce ----------------------------------------------------------------------------

;===================
;This happens when you put shit on top of shit:
;https://www.autoitscript.com/wiki/FAQ#Keys_virtually_stuck
;https://www.autoitscript.com/forum/topic/106408-shift-stuck-solved/
;Just a little note here: The issue is known for 2017-2009 = 8 years by now...
;===================
$dll = DllOpen("C:\Windows\System32\user32.dll")
Global Const $keys[8] = [0xa0, 0xa1, 0xa2, 0xa3, 0xa4, 0xa5, 0x5b, 0x5c]
;0xa0   LSHIFT
;0xa1   RSHIFT
;0xa2   LCTRL
;0xa3   RCTRL
;0xa4   LALT
;0xa5   RALT
;0x5b   LWIN
;0x5c   RWIN
Func UnstickKeys()
    For $vkvalue in $keys
        DllCall($dll,"int","keybd_event","int",$vkvalue,"int",0,"long",2,"long",0) ;Release each key
    Next
EndFunc

;===================
;functions:
;===================
Func pixelHasColor($handle, $x, $y, $colorString)
  if $handle = 0 Then
    MsgBox(64, "Error", "pixelHasColor needs a handle to work...")
    Exit -1
  EndIf

  If PixelGetColor($x, $y, $handle) = Dec($colorString, 2) Then
    Return True
  Else
    Return False
  EndIf
EndFunc

Func isWindowMinimized($handle)
  Local $iState = WinGetState($handle)
  If BitAND($iState,16) Then
    return True
  Else
    return False
  EndIf
EndFunc

Func isWindowMaximized($handle)
  Local $iState = WinGetState($handle)
  If BitAND($iState,32) Then
    return True
  Else
    return False
  EndIf
EndFunc

;===================
;script begin
;===================
;Parse command line
Opt("PixelCoordMode", 0) ; pixelFunctions coordinate sys relative to windows..
$numParams = $CmdLine[0]
$desiredFile = $CmdLine[1]
$exportFile = $CmdLine[2]
$startCommand = '"C:\Program Files (x86)\Microsoft Office\root\Office16\ONENOTE.exe" "' & $desiredFile & '"'
Run($startCommand)

;Get focus... Start by getting the correct Window-Handle...
;Wait until WinGetText returns a handle containing the text "Ribbon"
;This means that we've got the correct handle and not the handle of the
;Loading-Screen for example.
While 0 <> 1
  $titleRegex = "[REGEXPTITLE:(.*OneNote*)]"
  ; Tell autoit to search for a window (and activate it) containing "Ribbon":
  $handle = WinActivate($titleRegex, "Ribbon")
  If $handle <> 0 Then
    ExitLoop
 EndIf
 Sleep(100)
WEnd

;Maximize Window
While 0 <> 1
  WinSetState($handle, "", @SW_MAXIMIZE)
  $worked = isWindowMaximized($handle)
  If $worked Then
    ExitLoop
  EndIf
  Sleep(100)
WEnd

;Send("{LALT DOWN}{d}{c}{a}{p}{LALT UP}"):
UnstickKeys()
While 0 <> 1
  Send("{LALT DOWN}")
  Sleep(100)
  ; Wait for OneNote to draw the hint letters which are usually drawn when pressing the left ALT-Key.
  ; If it draws them it's ready for input... (Only way I found to really determine that..)
  $worked = pixelHasColor($handle, 151, 71, "525252")
  If $worked Then
    ExitLoop
  EndIf
WEnd
Send("{d}{c}{a}{p}{LALT UP}")
UnstickKeys()

;Save as dialogue will appear now.. Wait for it:
$saveAsHandle = WinWaitActive("Speichern unter", "", 5) ; max 5s
If $saveAsHandle = 0 Then
  MsgBox(64, "Error", "Save as dialogue did not appear.")
  Exit
EndIf
ControlSetText($saveAsHandle, "", "[CLASS:Edit; INSTANCE:1]", $exportFile)
ControlSend($saveAsHandle, "", "[CLASS:Edit; INSTANCE:1]", "{ENTER}")

;Minimize OneNote and wait for it to minimize (it will do as soon as saving is done...)
While 0 <> 1
  WinSetState($handle, "", @SW_MINIMIZE)
  $worked = isWindowMinimized($handle)
  If $worked Then
    ExitLoop
  EndIf
  Sleep(100)
WEnd

