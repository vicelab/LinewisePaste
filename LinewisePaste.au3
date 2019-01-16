; This script pastes either by stripping out tabs or by typing out the clipboard contents (replacing line feeds with down arrows)
; Written by A. Anderson using AutoIt


#include <MsgBoxConstants.au3>
#include <StringConstants.au3>

Global $g_bPaused = False
Global $makeanException = False
$dll = DllOpen("user32.dll")
AutoItSetOption( "WinTitleMatchMode", 2)

;HotKeySet("{Pause}", "TerminateLP") ; Press Pause/Break to terminate script
HotKeySet("!z", "Lpaste") ; Alt-z
HotKeySet("!v", "Tpaste") ; Alt-v

MsgBox($MB_SYSTEMMODAL, "LinewisePaste", "Use ALT-V to paste stripping out TABs." & @CRLF & @CRLF & "Use ALT-Z to paste a clipboard as individual letters," & @CRLF & "replacing line feeds w/ DOWN keys." & @CRLF & @CRLF & "To close the program, right-click the icon in the system tray.")

While 1
   Sleep(10000)
WEnd


;Func TerminateLP() ; This terminates the script upon pressing the hotkey defined above
;  Exit
;EndFunc   ;==>Terminate


Func Lpaste()
   $s = ClipGet()
   Sleep(700)
   $sSplit = StringSplit($s, @CRLF, $STR_ENTIRESPLIT)

   For $i = 1 To UBound($sSplit)-1
	  Sleep(200)
	  $line = StringSplit($sSplit[$i], "")
	  For $j =  1 To UBound($line)-1
		 Send($line[$j])
	  Next
	  Sleep(200)
	  Send("{DOWN}")
   Next
EndFunc   ;==>Lpaste


Func Tpaste()
   $s = ClipGet()
   $sSplit = StringSplit($s, @TAB)
   $t = ""
   For $i = 1 To UBound($sSplit)-1
	  $t = $t & $sSplit[$i]
   Next
   ClipPut($t)
   Sleep(700)
   Send("^v")
   ClipPut($s)
EndFunc   ;==>Tpaste



