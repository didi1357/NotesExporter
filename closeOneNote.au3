#cs ----------------------------------------------------------------------------

 AutoIt Version: 3.3.14.2
 Copyright:      Licenced under the terms of GPL v3 or later version
 Author:         Dietmar Malli

 Script Function:
	Simple Script to just close any open OneNote instance.

#ce ----------------------------------------------------------------------------

WinClose("[REGEXPTITLE:(.*OneNote*)]","Ribbon")

