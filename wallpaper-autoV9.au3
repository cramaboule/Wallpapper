#Region ;**** Directives created by AutoIt3Wrapper_GUI ****
#AutoIt3Wrapper_Icon=AutoItv11.ico
#AutoIt3Wrapper_Res_Fileversion=9.0.0.1
#AutoIt3Wrapper_Run_Tidy=y
#AutoIt3Wrapper_Run_Au3Stripper=y
#Au3Stripper_Parameters=/rsln
#EndRegion ;**** Directives created by AutoIt3Wrapper_GUI ****
#include <GDIPlus.au3>
#include <GUIConstantsEx.au3>
#include <WindowsConstants.au3>
#include <ButtonConstants.au3>
#include <ScreenCapture.au3>
#include <string.au3>
#include <INet.au3>
#include <File.au3>
#include <Array.au3>
#include <Date.au3>

Global $debug = 0

Global Const $head = "Wallpapaper V9"
; C:\Users\ma\AppData\Local\Wallpapper
Global Const $path = @LocalAppDataDir & "\Wallpapper\"
Global Const $pathjpg = $path & "Avion.jpg"
Global Const $pathjpg1 = $path & "Avion1.jpg"
Global Const $pathlog = $path & "Wallpapper.log"
Global Const $pathpg = $path & "page.txt"
Global Const $pathsmall = $path & "small.jpg"
Global Const $pathini = $path & "Avion.ini"
Global $Gui = 0, $pic, $tier, $demitier, $GUI_Button_Next, $GUI_Button_Go, $GUI_Button_Close, $z, $bar
Global Const $DesktopWidth = @DesktopWidth, $DesktopHeight = @DesktopHeight

DirCreate($path)

If FileExists($pathlog) Then
	$now = _NowCalc()
	$aTime = FileGetTime($pathlog, 1)
	$difference = _DateDiff('D', $aTime[0] & '/' & $aTime[1] & '/' & $aTime[2], $now) ;YYYY/MM/DD[ HH:MM:SS
	If $difference > 30 Then
		FileDelete($pathlog)
		$hLogFile = FileOpen($pathlog, 8 + 2)
	Else
		$hLogFile = FileOpen($pathlog, 1)
	EndIf
Else
	$hLogFile = FileOpen($pathlog, 8 + 2)
EndIf
_FileWriteLog($hLogFile, "---Start prg---")

Opt('WINTITLEMATCHMODE', 4)
Do
	$bar = WinGetPos("[CLASS:Shell_TrayWnd]")

Until IsArray($bar)

If Not FileExists($pathini) Then
	IniWrite($pathini, "Date", "Date", "2000/01/01")
	;IniWrite($pathini, "Date", "Date", _NowCalcDate())
EndIf

If $CmdLine[0] And $CmdLine[1] = "-a" Then
	$Lastwallpaper = IniRead($pathini, "Date", "Date", "2000/01/01")
	If $Lastwallpaper = _NowCalcDate() Then
		_FileWriteLog($hLogFile, "---END---")
		FileClose($hLogFile)
		Exit
	EndIf
EndIf

;TrayTip ( "Wallpaper", "Working...", 30)


$Links = _GetLinks()
_FileWriteLog($hLogFile, "---Links okay---")
_FileWriteLog($hLogFile, "---UBound=" & UBound($Links))
For $i = 0 To (UBound($Links) - 1)
	_FileWriteLog($hLogFile, "---$i=" & $i & " - " & StringReplace(StringStripWS($Links[$i], 7), Chr(10), "") & " ---")
Next
_FileWriteLog($hLogFile, "---_CreateGUI---")
_CreateGUI()
; [0]=ubound
; [1]=Source small
; [2]=$SmallWidth
; [3]=$Smallheight
; [4]=Source Big
; [5]=txt source
; [6]=Txt (Airline & Aircraft)
; [7]=
; [8]=

While 1
	$msg = GUIGetMsg()
	Select
		Case $msg = $GUI_EVENT_CLOSE Or $msg = $GUI_Button_Close
			_TrayBoxAnimate($Gui, 8)
			_FileWriteLog($hLogFile, "---Closed by user ---")
			Exit
		Case $msg = $GUI_Button_Next
			_TrayBoxAnimate($Gui, 8)
			_FileWriteLog($hLogFile, "---Changed by user ---")
			_FileWriteLog($hLogFile, "---_GetLinks inside while ---")
			$Links = _GetLinks()
			_FileWriteLog($hLogFile, "---UBound inside while =" & UBound($Links))
			For $i = 0 To (UBound($Links) - 1)
				_FileWriteLog($hLogFile, "---$i=" & $i & " - " & StringReplace(StringStripWS($Links[$i], 7), Chr(10), "") & " ---")
			Next
			_FileWriteLog($hLogFile, "---_CreateGUI inside while ---")
			_CreateGUI()
			_TrayBoxAnimate($Gui, 7)
		Case $msg = $GUI_Button_Go
			_FileWriteLog($hLogFile, "---Set by user ---")
			ExitLoop
	EndSelect
WEnd

_TrayBoxAnimate($Gui, 8)
GUIDelete($Gui)

_FileWriteLog($hLogFile, "---download pic ---")
InetGet($Links[4], $pathjpg)

;resize the picture
_FileWriteLog($hLogFile, "---_GDIPlus_Startup ---")
_GDIPlus_Startup()
$hImage2 = _GDIPlus_ImageLoadFromFile($pathjpg)
$width = _GDIPlus_ImageGetWidth($hImage2)
$Height = _GDIPlus_ImageGetHeight($hImage2)
$FDesktop = @DesktopHeight / @DesktopWidth
$Fact = 1
If $width > @DesktopWidth And $FDesktop > ($Height / $width) Then
	$Fact = @DesktopWidth / $width
ElseIf $Height > @DesktopHeight Then
	$Fact = @DesktopHeight / $Height
EndIf
$H1 = Round(($Fact * $Height), 0)
$W1 = Round(($Fact * $width), 0)
;
$hWnd = _WinAPI_GetDesktopWindow()
$hDC = _WinAPI_GetDC($hWnd)
$hBMP = _WinAPI_CreateCompatibleBitmap($hDC, $W1, $H1)
$hImage1 = _GDIPlus_BitmapCreateFromHBITMAP($hBMP)
$hGraphic = _GDIPlus_ImageGetGraphicsContext($hImage1)
_GDIPlus_GraphicsDrawImageRect($hGraphic, $hImage2, 0, 0, $W1, $H1)
;
$hBrush1 = _GDIPlus_BrushCreateSolid(0xFFFFFFFF) ; text color
$hBrush = _GDIPlus_BrushCreateSolid("0x60ffffff") ; layout
_GDIPlus_GraphicsFillRect($hGraphic, 0, 0, $W1, 20, $hBrush)
_GDIPlus_GraphicsDrawString($hGraphic, $Links[6], 0, 0, "Arial Black")
;
;~ $CLSID = _GDIPlus_EncodersGetCLSID("bmp")
;~ _GDIPlus_ImageSaveToFileEx($hImage1, $pathbmp, $CLSID)
;
$TParam = _GDIPlus_ParamInit(1)
$Datas = DllStructCreate("int Quality")
DllStructSetData($Datas, "Quality", '100')
_GDIPlus_ParamAdd($TParam, $GDIP_EPGQUALITY, 1, $GDIP_EPTLONG, DllStructGetPtr($Datas))
$Param = DllStructGetPtr($TParam)
;
$CLSID = _GDIPlus_EncodersGetCLSID("JPG")
_GDIPlus_ImageSaveToFileEx($hImage1, $pathjpg1, $CLSID, $Param)
;
_WinAPI_DeleteObject($hBMP)
_WinAPI_ReleaseDC($hWnd, $hDC)
_GDIPlus_ImageDispose($hImage1)
_GDIPlus_ImageDispose($hImage2)
_GDIPlus_GraphicsDispose($hGraphic)
_GDIPlus_Shutdown()
_ChangeDesktopWallpaper($pathjpg1, 0)

IniWrite($pathini, "Date", "Date", _NowCalcDate())
;~ Sleep(5000) ; let some time to read the tray tip !!! :-)

;~ TrayTip ( "","", Default)
;~ TrayTip ( "Etape4", "Quit", 5)
;~ Sleep(500)
_FileWriteLog($hLogFile, "---END---")
FileClose($hLogFile)
Exit
;~ ===================================================================================================FUNC===========================================
Func _CreateGUI()
	If $Gui <> 0 Then GUIDelete($Gui)
	If $Links[2] < 150 Then
		$width = 150 ;need space for the buttons
		$p = 0
	Else
		$width = $Links[2]
		$p = 1
	EndIf
	;
	_FileWriteLog($hLogFile, "---$head=" & $head & " ---")
	_FileWriteLog($hLogFile, "---$width=" & $width & " ---")
	_FileWriteLog($hLogFile, "---$Links[3]=" & $Links[3] & " ---")
	_FileWriteLog($hLogFile, "---$DesktopWidth=" & $DesktopWidth & " ---")
	_FileWriteLog($hLogFile, "---$DesktopHeight=" & $DesktopHeight & " ---")
	_FileWriteLog($hLogFile, "---$bar[3]=" & $bar[3] & " ---")
	_FileWriteLog($hLogFile, "---$WS_POPUP=" & $WS_POPUP & " ---")
	_FileWriteLog($hLogFile, "---$WS_BORDER=" & $WS_BORDER & " ---")
	;
	If @OSVersion = "WIN_7" Or @OSVersion = "WIN_VISTA" Then
		$Gui = GUICreate($head, $width - 2, $Links[3] + 40, $DesktopWidth - $width, $DesktopHeight - ($Links[3] + 40 + $bar[3]), BitOR($WS_POPUP, $WS_BORDER))
	ElseIf @OSVersion = "WIN_XP" Or @OSVersion = "WIN_XPe" Then
		$Gui = GUICreate($head, $width - 2, $Links[3] + 40, $DesktopWidth - $width, $DesktopHeight - ($Links[3] + $bar[3] + 48), BitOR($WS_POPUP, $WS_BORDER))
	ElseIf @OSVersion = "WIN_8" Or @OSVersion = "WIN_81" Then
		$Gui = GUICreate($head, $width - 2, $Links[3] + 40, $DesktopWidth - $width, $DesktopHeight - ($Links[3] + $bar[3] + 41), BitOR($WS_POPUP, $WS_BORDER))
	Else
		$Gui = GUICreate($head, $width - 2, $Links[3] + 40, $DesktopWidth - $width, $DesktopHeight - ($Links[3] + 40 + $bar[3]), BitOR($WS_POPUP, $WS_BORDER))
	EndIf
	$pic = GUICtrlCreatePic($pathsmall, (($width - $Links[2]) / 2) - $p, -1, $Links[2], $Links[3], $WS_BORDER)
	$tier = $width / 3
	$demitier = $tier / 2
	$GUI_Button_Next = GUICtrlCreateButton("Change", $demitier - 25, $Links[3] + 5, 50, 30, BitOR($BS_DEFPUSHBUTTON, $BS_FLAT))
	$GUI_Button_Go = GUICtrlCreateButton("Set", $demitier + $tier - 25, $Links[3] + 5, 50, 30, $BS_FLAT)
	$GUI_Button_Close = GUICtrlCreateButton("Close", $demitier + $tier + $tier - 25, $Links[3] + 5, 50, 30, $BS_FLAT)
	TrayTip("", "", Default)
	_TrayBoxAnimate($Gui, 7)
	GUISetState()
EndFunc   ;==>_CreateGUI

Func _ChangeDesktopWallpaper($bmp, $style = 0)
	;===============================================================================
	;
	; Function Name:    _ChangeDesktopWallPaper
	; Description:       Update WallPaper Settings
	;Usage:              _ChangeDesktopWallPaper(@WindowsDir & '\' & 'zapotec.bmp',1)
	; Parameter(s):     $bmp - Full Path to BitMap File (*.bmp)
	;                              [$style] - 0 = Centered, 1 = Tiled, 2 = Stretched
	; Requirement(s):   None.
	; Return Value(s):  On Success - Returns 0
	;                   On Failure -   -1
	; Author(s):        FlyingBoz
	;Thanks:        Larry - DllCall Example - Tested and Working under XPHome and W2K Pro
	;                     Excalibur - Reawakening my interest in Getting This done.
	;
	;===============================================================================

	If Not FileExists($bmp) Then Return -1
	;The $SPI*  values could be defined elsewhere via #include - if you conflict,
	; remove these, or add if Not IsDeclared "SPI_SETDESKWALLPAPER" Logic
	Local $SPI_SETDESKWALLPAPER = 20
	Local $SPIF_UPDATEINIFILE = 1
	Local $SPIF_SENDCHANGE = 2
	Local $REG_DESKTOP = "HKEY_CURRENT_USER\Control Panel\Desktop"
	If $style = 1 Then
		RegWrite($REG_DESKTOP, "TileWallPaper", "REG_SZ", 1)
		RegWrite($REG_DESKTOP, "WallpaperStyle", "REG_SZ", 0)
	Else
		RegWrite($REG_DESKTOP, "TileWallPaper", "REG_SZ", 0)
		RegWrite($REG_DESKTOP, "WallpaperStyle", "REG_SZ", $style)
	EndIf


	DllCall("user32.dll", "int", "SystemParametersInfo", _
			"int", $SPI_SETDESKWALLPAPER, _
			"int", 0, _
			"str", $bmp, _
			"int", BitOR($SPIF_UPDATEINIFILE, $SPIF_SENDCHANGE))
	Return 0
EndFunc   ;==>_ChangeDesktopWallpaper

Func _TrayBoxAnimate($TBGui, $Xstyle = 1, $Xspeed = 1500)
	; $Xstyle - 1=Fade, 3=Explode, 5=L-Slide, 7=R-Slide, 9=T-Slide, 11=B-Slide,
	;13=TL-Diag-Slide, 15=TR-Diag-Slide, 17=BL-Diag-Slide, 19=BR-Diag-Slide
	Local $Xpick = StringSplit('80000,90000,40010,50010,40001,50002,40002,50001,40004,50008,40008,50004,40005,5000a,40006,50009,40009,50006,4000a,50005', ",")
	DllCall("user32.dll", "int", "AnimateWindow", "hwnd", $TBGui, "int", $Xspeed, "long", "0x000" & $Xpick[$Xstyle])
EndFunc   ;==>_TrayBoxAnimate

Func _GetLinks()
	Do ;3878175
		Local $rdm = Random(1, 7000000, 1)
		$small = "-6.jpg"
		$big = "-12.jpg"
;~ 		Local $rdm = 2062716 ;makes problem
;~ 		Local $rdm = 1717766
;~ 		Local $rdm = 975633 ; is making problems
;~ http://www.airliners.net/photo/219906
;~ http://www.airliners.net/photo/Airbus/Airbus-A340-642/219906
;~ http://cdn-www.airliners.net/photos/airliners/6/0/9/0219906.jpg?v=v20
;~ http://imgproc.airliners.net/photos/airliners/6/0/9/0219906-v20-15.jpg (moyenne)
;~ http://imgproc.airliners.net/photos/airliners/6/0/9/0219906-v20-6.jpg (petite)
;~ http://imgproc.airliners.net/photos/airliners/6/0/9/0219906-v20-12.jpg (grande)
		If $debug Then ConsoleWrite($rdm & @CRLF)
		_FileWriteLog($hLogFile, "---$rdm=" & $rdm & "---")
		Local $flag = 0
		Local $Text = _INetGetSource("http://www.airliners.net/photo/" & $rdm, 1)
		If @error Then
			$flag = 1
		EndIf
		If $debug Then ConsoleWrite($Text & @CRLF)
		If Not ($flag) Then
			$aText = _StringBetween($Text, '<html', '</html>')
			If @error Then
				$flag = 1
			EndIf
		EndIf
		If Not ($flag) Then
			$aaaText = _StringBetween($aText[0], '<div class="pdp-image-wrapper">', '</a>')
			If @error Then
				$flag = 1
			EndIf
		EndIf
		If Not ($flag) Then
			Local $Href = _StringBetween($aaaText[0], '<img src="', '"')
			If @error Then
				$flag = 1
			EndIf
		EndIf
	Until $flag = 0
	;ConsoleWrite($Href[0] & @CRLF)
	;http://cdn-www.airliners.net/photos/airliners/8/3/3/2448338.jpg?v=v20
	;http://imgproc.airliners.net/photos/airliners/8/3/3/2448338-v20-15.jpg
	;http://imgproc.airliners.net/photos/airliners/8/3/3/2448338-v20-6.jpg
	$Href[0] = StringReplace($Href[0], "cdn-www.airliners.net", "imgproc.airliners.net")
	$tt = StringInStr($Href[0], ".jpg?v=")
	$sID = StringMid($Href[0], $tt + 7, StringLen($Href[0]))
	;ConsoleWrite($sID & @CRLF)
	$Href[0] = StringReplace($Href[0], ".jpg?v=" & $sID, "-" & $sID & $small)
	;ConsoleWrite($Href[0] & @CRLF)
	ReDim $Href[UBound($Href) + 7]
	For $i = UBound($Href) - 2 To 0 Step -1 ; décale toute la 'array' de +1
		$Href[$i + 1] = $Href[$i]
	Next
	$Href[0] = UBound($Href) - 1
	$Href[5] = $aText[0]
	;ConsoleWrite($Href[5] & @CRLF)
	InetGet($Href[1], $pathsmall)
	$Textpg = _StringBetween($Href[5], "<title>", "</title>")
	$Href[6] = $Textpg[0]
	$Href[4] = StringReplace($Href[1], $small, $big)
	_GDIPlus_Startup()
	$hImage6 = _GDIPlus_ImageLoadFromFile($pathsmall)
	$Href[2] = _GDIPlus_ImageGetWidth($hImage6)
	$Href[3] = _GDIPlus_ImageGetHeight($hImage6)
	_GDIPlus_ImageDispose($hImage6)
	_GDIPlus_Shutdown()
	;ConsoleWrite($pathsmall & @CRLF)
	;_ArrayDisplay($Href)
	Return $Href
EndFunc   ;==>_GetLinks
; [0]=ubound
; [1]=Source smaill
; [2]=$SmallWidth
; [3]=$Smallheight
; [4]=Source Big
; [5]=txt source
; [6]=Txt (Airline & Aircraft)
; [7]=
; [8]=
