; This script pastes either by stripping out tabs or by typing out the clipboard contents (replacing line feeds with down arrows)
; This can be useful when constructing lists of file names and entering them into an ArcMap Batch window
; Written by A. Anderson using AutoIt


#include <MsgBoxConstants.au3>
#include <StringConstants.au3>
#include <Misc.au3>

Local $hDLL = DllOpen("user32.dll")

Global $g_bPaused = False
Global $makeanException = False
$dll = DllOpen("user32.dll")

AutoItSetOption( "WinTitleMatchMode", 2)

HotKeySet("!d", "Dpaste") ; Alt-d
HotKeySet("!z", "Lpaste") ; Alt-z
HotKeySet("!v", "Tpaste") ; Alt-v
HotKeySet("{PAUSE}", "stopThePresses")

MsgBox($MB_SYSTEMMODAL, "LinewisePaste                                           rev. 2018-02-06", "Use ALT-V to paste stripping out TABs." & @CRLF & @CRLF & "Use ALT-Z to type out the clipboard as individual letters," & @CRLF & "replacing line feeds with DOWN keys." & @CRLF & @CRLF & "Pressing PAUSE will halt a typing Paste." & @CRLF & @CRLF & "To close the program, right-click the icon in the system tray.")
$goOn = True


While 1
   Sleep(10000)
WEnd


;Func TerminateLP() ; This terminates the script upon pressing the hotkey defined above
;  Exit
;EndFunc   ;==>Terminate

Func stopThePresses()
   $goOn = False
EndFunc ;==>stopThePresses

Func Dpaste()
   Local $hDLL = DllOpen("user32.dll")
   $s = ClipGet()
   Sleep(700)
   $sSplit = StringSplit($s, @CRLF, $STR_ENTIRESPLIT)

   For $i = 1 To UBound($sSplit)-1
	  ;Sleep(200)
	  $line = StringSplit($sSplit[$i], "")
	  For $j =  1 To UBound($line)-1
		 Send($line[$j])
		 if $goOn == False Then
			ExitLoop(2)
		 EndIf
	  Next
	  Send("{DOWN}")
	  While 1
		 If _IsPressed("2D", $hDLL) Then
			; Wait until key is released.
			While _IsPressed("2D", $hDLL)
			   Sleep(50)
			WEnd
			ExitLoop
         EndIf
		 Sleep(50)
	  WEnd
   Next
   $goOn = True
   DllClose($hDLL)
EndFunc   ;==>Dpaste

Func Lpaste()
   $s = ClipGet()
   Sleep(700)
   $sSplit = StringSplit($s, @CRLF, $STR_ENTIRESPLIT)

   For $i = 1 To UBound($sSplit)-1
	  Sleep(200)
	  $line = StringSplit($sSplit[$i], "")
	  For $j =  1 To UBound($line)-1
		 Send($line[$j])
		 if $goOn == False Then
			ExitLoop(2)
		 EndIf
	  Next
	  Send("{DOWN}")
	  Sleep(200)
   Next
   $goOn = True
EndFunc   ;==>Lpaste


Func Tpaste()
   $s = ClipGet()
   $sSplit = StringSplit($s, @TAB)
   $t = ""
   For $i = 1 To UBound($sSplit)-1
	  $t = $t & $sSplit[$i]
   Next
   ClipPut($t)
   Sleep(500)
   Send("^v")
   Sleep(500)
   ClipPut($s)
EndFunc   ;==>Tpaste




