#Region ;**** Directives created by AutoIt3Wrapper_GUI ****
#AutoIt3Wrapper_Icon=..\AutoItv11.ico
#AutoIt3Wrapper_Res_Comment=Convert and resize from/to *NEW* WEBP, JPG, BMP, GIF, PNG,...
#AutoIt3Wrapper_Res_Description=Convert and resize from/to *NEW* WEBP, JPG, BMP, GIF, PNG,...
#AutoIt3Wrapper_Res_Fileversion=3.0.2.0
#AutoIt3Wrapper_Res_ProductName=Pics Converter V3
#AutoIt3Wrapper_Res_ProductVersion=3.0.2.0
#AutoIt3Wrapper_Res_CompanyName=cramaboule.com
#AutoIt3Wrapper_Run_Before=%scriptdir%\..\WriteTimestampAndVersion.exe "%in%"
#AutoIt3Wrapper_Run_After=copy %in% ..\..\Github\PicsConverter\
#AutoIt3Wrapper_Run_After=copy %out% ..\..\Github\PicsConverter\
#AutoIt3Wrapper_Run_After=copy ExtMsgBox.au3 ..\..\Github\PicsConverter\
#AutoIt3Wrapper_Run_After=copy StringSize.au3 ..\..\Github\PicsConverter\
#AutoIt3Wrapper_Run_Tidy=y
#Tidy_Parameters=/reel
#AutoIt3Wrapper_Run_Au3Stripper=y
#Au3Stripper_Parameters=/mo
#EndRegion ;**** Directives created by AutoIt3Wrapper_GUI ****
#Region    ;Timestamp =====================
#    Last compile at : 2025/12/09 10:23:11
#EndRegion ;Timestamp =====================
#cs ----------------------------------------------------------------------------

	AutoIt Version: 3.3.18.0
	Author:         Cramaboule
	Date:			October 2009 V1

	Script Function: 	'Pics Converter V3 in (almoast) Pure AutoIt' made by Cramaboule Mai 2023
						Thanks to AdmiralAlkex for his help ! (on V1)

						Convert from/to JPG, BMP, GIF, PNG ,...!!! AND NEW WEBP
						Enjoy !

	Link: WebP: https://developers.google.com/speed/webp/download  /  https://developers.google.com/speed/webp/docs/cwebp

	Bug:

	To Do:	how to keep metadata?

	V3.0.2.0	04.12.2025:
				New: New version of WEBP
				Changed: path to WEBP
				New: Resize from/to WEBP
				Changed: rearrange Guis
	V3.0.1.1	04.06.2023:
				fix bugs
				Added: pourcent
	V3.0.1.0	02.06.2023:
				Improved: faster search with _ArrayConcatenate()
				Changed: -lossless can be used with -q (for WebP)
				Improved: faster conversion using -mt for WebP
					https://developers.google.com/speed/webp/docs/cwebp
				Improved: optimize decodig webp
					https://github.com/webmproject/libwebp/blob/0905f61c8511f080bec75ba98f67d53bb2906ccf/doc/tools.md
				Added: Decoders from GDI+
				Added: auto select input encoder

	V3.0.0.0	31.05.2023:
				Added: WebP
				Added: Subfolder
				Changed: rearange Gui
				Remove: small bug

	V2.0.0.0	07.07.2023:
				Resizing fonction

	V1.2 bugs fixed
	V1.1 added new features
	V1.0 first realese

#ce ----------------------------------------------------------------------------
#include <ScreenCapture.au3>
#include <GDIPlus.au3>
#include <File.au3>
#include <ProgressConstants.au3>
#include <ButtonConstants.au3>
#include <ComboConstants.au3>
#include <EditConstants.au3>
#include <GUIConstantsEx.au3>
#include <SliderConstants.au3>
#include <StaticConstants.au3>
#include <WindowsConstants.au3>
#include <StringConstants.au3>
#include <FileConstants.au3>
#include <Array.au3>
#include 'ExtMsgBox.au3'

$sVersion = 'V3.0.2.0'
$head = 'Pics Conversion ' & $sVersion

Local $Param = 0, $Decoder, $ToCombo, $ToComboOut, $OldOutEncoder, $Oldpxpercent, $Label2
Local $OldValSlider = '0', $OldJPGQuality = '100', $OldHeight, $Oldwidth, $OldCheckRatio, $OldLossless, $OldResize, $Parameter, $WidthHeight[2]
Dim $aInterpolation[2][7] = [[$GDIP_INTERPOLATIONMODE_HIGHQUALITYBICUBIC, $GDIP_INTERPOLATIONMODE_HIGHQUALITYBILINEAR, $GDIP_INTERPOLATIONMODE_NEARESTNEIGHBOR, $GDIP_INTERPOLATIONMODE_BICUBIC, $GDIP_INTERPOLATIONMODE_BILINEAR, $GDIP_INTERPOLATIONMODE_HIGHQUALITY, $GDIP_INTERPOLATIONMODE_LOWQUALITY], ['Bicubic HQ (default)', 'Nearest neighbor', 'Bilinear HQ', 'Bicubic (low)', 'Bilinear (low)', 'High-quality', 'Low-quality']]
Global $pathWebP = _CheckWebP()

_GDIPlus_Startup()
$testBMP = _ScreenCapture_Capture("", 0, 0, 1, 1)
$hImage = _GDIPlus_BitmapCreateFromHBITMAP($testBMP)
$Decoder = _GDIPlus_Decoders()
$Encoder = _GDIPlus_Encoders()
_GDIPlus_ImageDispose($hImage)
_WinAPI_DeleteObject($testBMP)
_GDIPlus_Shutdown()

If $pathWebP <> '' Then
	$ToCombo &= 'WEBP|'
	$ToComboOut &= 'WEBP|'
EndIf

For $i = 1 To $Encoder[0][0]
	$Split = StringSplit($Encoder[$i][6], ";")
	For $j = 1 To $Split[0]
		$ToComboOut &= StringTrimLeft($Split[$j], 2) & "|"
	Next
Next

For $i = 1 To $Decoder[0][0]
	$Split = StringSplit($Decoder[$i][6], ";")
	For $j = 1 To $Split[0]
		$ToCombo &= StringTrimLeft($Split[$j], 2) & "|"
	Next
Next

$Conv = GUICreate($head, 570, 210, -1, -1)
$Group1 = GUICtrlCreateGroup(" Input ", 5, 5, 140, 145)
$InputEncoder = GUICtrlCreateCombo("", 15, 120, 120, 25)
GUICtrlSetData(-1, $ToCombo)
$InputFolder = GUICtrlCreateInput("Input Folder", 15, 25, 120, 21)
$BrowseInput = GUICtrlCreateButton("Browse...", 60, 50, 75, 25, $WS_GROUP)
$Subfolder = GUICtrlCreateCheckbox("Subfolers ?", 60, 77, 75, 21)
$Label9 = GUICtrlCreateLabel("Convert from:", 15, 100, 67, 17)
GUICtrlCreateGroup("", -99, -99, 1, 1)

$Group4 = GUICtrlCreateGroup(" Resize ", 155, 5, 160, 145)
$Resizing = GUICtrlCreateCheckbox("Resizing", 165, 25, 70, 18)
$pxpercent = GUICtrlCreateCombo('px', 250, 25, 40, 25)
GUICtrlSetData(-1, '%')
$Ratio = GUICtrlCreateCheckbox("Keep aspect ratio", 175, 47, 115, 20)
GUICtrlSetState(-1, $GUI_CHECKED)

$Width = GUICtrlCreateInput("", 170, 72, 49, 21)
$Label4 = GUICtrlCreateLabel("px", 220, 77, 15, 17)
$Height = GUICtrlCreateInput("", 248, 72, 49, 21)
$Label5 = GUICtrlCreateLabel("px", 298, 77, 15, 17)
$Label6 = GUICtrlCreateLabel("w:", 158, 77, 10, 17)
$Label7 = GUICtrlCreateLabel("h:", 238, 77, 10, 17)

$Label3 = GUICtrlCreateLabel("Interpolation mode:", 168, 100, 104, 17)
$Interpolation = GUICtrlCreateCombo($aInterpolation[1][0], 165, 120, 120, 25, BitOR($CBS_DROPDOWN, $CBS_AUTOHSCROLL))
GUICtrlSetData(-1, $aInterpolation[1][1] & "|" & $aInterpolation[1][2] & "|" & $aInterpolation[1][3] & "|" & $aInterpolation[1][4] & "|" & $aInterpolation[1][5] & "|" & $aInterpolation[1][6])

GUICtrlSetState($Ratio, $GUI_DISABLE)
GUICtrlSetState($Width, $GUI_DISABLE)
GUICtrlSetState($Label4, $GUI_DISABLE)
GUICtrlSetState($Height, $GUI_DISABLE)
GUICtrlSetState($Label5, $GUI_DISABLE)
GUICtrlSetState($Label3, $GUI_DISABLE)
GUICtrlSetState($Label6, $GUI_DISABLE)
GUICtrlSetState($Label7, $GUI_DISABLE)
GUICtrlSetState($Interpolation, $GUI_DISABLE)
GUICtrlSetState($pxpercent, $GUI_DISABLE)

GUICtrlCreateGroup("", -99, -99, 1, 1)

$Group2 = GUICtrlCreateGroup(" Output ", 325, 5, 140, 145)
$OutputEncoder = GUICtrlCreateCombo("", 335, 120, 120, 25)
GUICtrlSetData(-1, $ToComboOut)
$OutputFolder = GUICtrlCreateInput("Output Folder", 335, 25, 120, 21)
$BrowseOutput = GUICtrlCreateButton("Browse...", 380, 50, 75, 25, $WS_GROUP)
$Label1 = GUICtrlCreateLabel("Convert to:", 335, 100, 56, 17)
GUICtrlCreateGroup("", -99, -99, 1, 1)

$Group3 = GUICtrlCreateGroup(" Quality ", 475, 5, 90, 145)
$Lossless = GUICtrlCreateCheckbox('Lossless', 480, 25, 60, 25)
$Slider = GUICtrlCreateSlider(515, 47, 35, 100, BitOR($TBS_VERT, $TBS_TOP, $TBS_LEFT))
$JPGQlty = GUICtrlCreateInput("100", 485, 87, 30, 21)

GUICtrlSetState($Group3, $GUI_ENABLE)
GUICtrlSetState($Slider, $GUI_DISABLE)
GUICtrlSetState($JPGQlty, $GUI_DISABLE)
GUICtrlSetState($Lossless, $GUI_DISABLE)

GUICtrlCreateGroup("", -99, -99, 1, 1)

$GO = GUICtrlCreateButton("Convert", 185, 160, 200, 40, $WS_GROUP)
GUISetState(@SW_SHOW)

While 1
	$OutEncoder = GUICtrlRead($OutputEncoder)
	$sPxpercent = GUICtrlRead($pxpercent)
	$ValSlider = GUICtrlRead($Slider)
	$JPGQuality = GUICtrlRead($JPGQlty)
	$iHeight = GUICtrlRead($Height)
	$iWidth = GUICtrlRead($Width)
	$nMsg = GUIGetMsg()
	Select
		Case $nMsg = $GUI_EVENT_CLOSE
			Exit
		Case $nMsg = $BrowseInput
			$sInFold = GUICtrlRead($InputFolder)
			If $sInFold = "Choose Folder" Then $sInFold = ""
			$sInFold = FileSelectFolder("Choose a folder", $sInFold, 7, '', GUICreate(""))
			If $sInFold <> "" Then
				$InFold = $sInFold
				GUICtrlSetData($InputFolder, 'Please wait...')
				$sFindExt = _FindExtention($InFold, $ToCombo)
				If $sFindExt <> '' Then
					GUICtrlSetData($InputEncoder, $sFindExt)
				EndIf
				GUICtrlSetData($InputFolder, $InFold)
				If Not (StringInStr(GUICtrlRead($OutputFolder), "\")) Then
					GUICtrlSetData($OutputFolder, $InFold)
				EndIf
			EndIf
		Case $nMsg = $BrowseOutput
			$sOutFold = GUICtrlRead($OutputFolder)
			If $sOutFold = "Choose Folder" Then $sOutFold = ""
			$sOutFold = FileSelectFolder("Choose a folder", $sOutFold, 1, '', GUICreate(""))
			If $sOutFold <> "" Then
				$OutFold = $sOutFold
				GUICtrlSetData($OutputFolder, $OutFold)
			EndIf
		Case $OutEncoder <> $OldOutEncoder
			If ($OutEncoder <> 'JPG') And ($OutEncoder <> 'WEBP') Then
				GUICtrlSetState($Lossless, $GUI_UNCHECKED)
				GUICtrlSetState($Lossless, $GUI_DISABLE)
				GUICtrlSetState($Group3, $GUI_DISABLE)
				GUICtrlSetState($Slider, $GUI_DISABLE)
				GUICtrlSetState($JPGQlty, $GUI_DISABLE)
				GUICtrlSetState($Lossless, $GUI_DISABLE)
				GUICtrlSetState($Resizing, $GUI_ENABLE)
			Else
				If $OutEncoder = 'JPG' Then
					GUICtrlSetState($Lossless, $GUI_UNCHECKED)
					GUICtrlSetState($Lossless, $GUI_DISABLE)
					GUICtrlSetState($Group3, $GUI_ENABLE)
					GUICtrlSetState($Slider, $GUI_ENABLE)
					GUICtrlSetState($JPGQlty, $GUI_ENABLE)
					GUICtrlSetState($Resizing, $GUI_ENABLE)
				ElseIf $OutEncoder = 'WEBP' Then
					GUICtrlSetState($Lossless, $GUI_ENABLE)
					GUICtrlSetState($Slider, $GUI_ENABLE)
					GUICtrlSetState($JPGQlty, $GUI_ENABLE)
				EndIf
			EndIf
			If _IsChecked($Resizing) <> $OldResize Then
				_CheckResize(_IsChecked($Resizing))
			EndIf
			$OldOutEncoder = $OutEncoder
		Case $ValSlider <> $OldValSlider
			GUICtrlSetData($JPGQlty, 100 - $ValSlider)
			$OldValSlider = $ValSlider
		Case $JPGQuality <> $OldJPGQuality
			$JPGQuality = _checkValue($JPGQuality)
			GUICtrlSetData($JPGQlty, $JPGQuality)
			GUICtrlSetData($Slider, 100 - $JPGQuality)
			$OldJPGQuality = $JPGQuality
		Case $nMsg = $Resizing
			_CheckResize(_IsChecked($Resizing))
			$OldResize = _CheckResize(_IsChecked($Resizing))
		Case $sPxpercent <> $Oldpxpercent
			If $sPxpercent = 'px' Then
				GUICtrlSetData($Label4, 'px')
				GUICtrlSetData($Label5, 'px')
			Else
				GUICtrlSetData($Label4, '%')
				GUICtrlSetData($Label5, '%')
				$sWidth = GUICtrlRead($Width)
				If $sWidth < 0 And $sWidth > 100 Then
					GUICtrlSetData($Width, '100')
				EndIf
				$sHeight = GUICtrlRead($Height)
				If $sHeight < 0 And $sHeight > 100 Then
					GUICtrlSetData($sHeight, '100')
				EndIf
			EndIf
			$Oldpxpercent = $sPxpercent
		Case (($iWidth <> $Oldwidth Or $iHeight <> $OldHeight) And _IsChecked($Ratio) And GUICtrlRead($pxpercent) = '%')
			If $iWidth <> $Oldwidth Then
				$iWidth = _checkValue($iWidth)
				$iHeight = $iWidth
			Else
				$iHeight = _checkValue($iHeight)
				$iWidth = $iHeight
			EndIf
			GUICtrlSetData($Width, $iHeight)
			GUICtrlSetData($Height, $iWidth)
			$Oldwidth = $iWidth
			$OldHeight = $iHeight
		Case (($iWidth <> $Oldwidth Or $iHeight <> $OldHeight) And Not (_IsChecked($Ratio)) And GUICtrlRead($pxpercent) = '%')
			If $iWidth <> $Oldwidth Then
				$iWidth = _checkValue($iWidth)
				GUICtrlSetData($Width, $iWidth)
			Else
				$iHeight = _checkValue($iHeight)
				GUICtrlSetData($Height, $iHeight)
			EndIf
			$Oldwidth = $iWidth
			$OldHeight = $iHeight
		Case _IsChecked($Ratio) <> $OldCheckRatio And GUICtrlRead($pxpercent) = '%'
			GUICtrlSetData($Height, $iWidth)
			$OldCheckRatio = _IsChecked($Ratio)
			$Oldwidth = $iWidth
			$OldHeight = $iHeight
			;	------------------------------------- Convert ----------------------------------------
		Case $nMsg = $GO
			$InPath = GUICtrlRead($InputFolder)
			$OutPath = GUICtrlRead($OutputFolder)
			$InEncoder = GUICtrlRead($InputEncoder)
			$OutEncoder = GUICtrlRead($OutputEncoder)
			If Not (StringInStr($InPath, "\")) Or Not (StringInStr($OutPath, "\")) Then
				_ExtMsgBox(16, 0, "Caution!", "Please select a folder!", 0, $Conv)
			ElseIf ($InEncoder = "") Or ($InEncoder = $OutEncoder And $OutEncoder <> "JPG" And $OutEncoder <> "WEBP" And Not (_IsChecked($Resizing))) Then
				_ExtMsgBox(16, 0, "Caution!", "Please choose different encoder/decoder or enable resizing", 0, $Conv)
			ElseIf GUICtrlRead($Width) = '' And GUICtrlRead($Height) = '' And _IsChecked($Resizing) Then
				_ExtMsgBox(16, 0, "Caution!", "Width and/or height cannot be empty", 0, $Conv)
			Else
;~ 				do the conversion process...
;~ 				do the progress bar GUI
				$aPos = WinGetPos($Conv)
				$iWinWidth = 550
				$iWinHeight = 135
				$Form1 = GUICreate("", $iWinWidth, $iWinHeight, ($aPos[0] + ($aPos[2] / 2)) - ($iWinWidth / 2), ($aPos[1] + ($aPos[3] / 2)) - ($iWinHeight / 2), BitOR($WS_POPUP, $WS_BORDER), $WS_EX_TOOLWINDOW, $Conv)
				Global $ProgFile = GUICtrlCreateProgress(10, 10, $iWinWidth - 20, 20, $PBS_SMOOTH)
				$Label2 = GUICtrlCreateLabel("", 10, 35, $iWinWidth - 20, 20, $SS_LEFT)
				$ProgAll = GUICtrlCreateProgress(10, 60, $iWinWidth - 20, 20, $PBS_SMOOTH)
				$Label3 = GUICtrlCreateLabel("", 10, 85, $iWinWidth - 20, 20, $SS_LEFT)
				$Label1 = GUICtrlCreateLabel("", 10, 110, $iWinWidth - 20, 20, $SS_LEFT)
				GUISetState(@SW_SHOW)
				If $OutEncoder = "JPG" Then ; Set JPG quality
					$TParam = _GDIPlus_ParamInit(1)
					$Datas = DllStructCreate("int Quality")
					DllStructSetData($Datas, "Quality", $JPGQuality)
					_GDIPlus_ParamAdd($TParam, $GDIP_EPGQUALITY, 1, $GDIP_EPTLONG, DllStructGetPtr($Datas))
					$Param = DllStructGetPtr($TParam)
				EndIf
				;
				If _IsChecked($Resizing) Then
					For $j = 0 To UBound($aInterpolation, 2) - 1 ; get the interpolation Mode according to the ComboBox
						If GUICtrlRead($Interpolation) = $aInterpolation[1][$j] Then
							$iInterpolation = $aInterpolation[0][$j]
						EndIf
					Next
				EndIf
;~ 				==================================================== process itself ==============================================
				Dim $FileList[1]
				$iFiles = _FindPathName($FileList, $InPath, "*." & $InEncoder, _IsChecked($Subfolder))
				If $iFiles <= 0 Then
					_ExtMsgBox(16, 0, "Caution!", "No files found or invalid path!", 0, $Conv)
				Else
					_GDIPlus_Startup()
					If $OutEncoder <> 'WEBP' Then
						$clsid = _GDIPlus_EncodersGetCLSID($OutEncoder)
					EndIf
					$nBin = 0
					If $OutEncoder = 'WEBP' Then $nBin = BitOR($nBin, 4)  ; cwebp : compresse un fichier image en fichier WebP
					If $InEncoder = 'WEBP' Then $nBin = BitOR($nBin, 2) ; dwebp : décompresser un fichier WebP dans un fichier image
					If _IsChecked($Resizing) Then $nBin = BitOR($nBin, 1)
;~ 					ConsoleWrite($nBin & @CRLF)
					For $i = 1 To $FileList[0]
						$iProgFile = 0
						GUICtrlSetData($ProgFile, $iProgFile)
						GUICtrlSetData($ProgAll, ($i / $FileList[0]) * 100)
						$sPicsIn = $FileList[$i]
						$sTempPath = StringReplace($sPicsIn, "." & $InEncoder, "." & $OutEncoder) ; change extention
						$sPicsOut = StringReplace($sTempPath, $InPath, $OutPath)
						GUICtrlSetData($Label2, $sPicsIn)
						GUICtrlSetData($Label3, $sPicsOut)
						GUICtrlSetData($Label1, $i & " / " & $FileList[0] & ' - ' & Round(($i / $FileList[0]) * 100, 1) & '%')
						$aPath = StringSplit($sPicsOut, '\')
						$sPathOut = StringReplace($sPicsOut, $aPath[UBound($aPath) - 1], '') ; out path (without file name)
						If Not FileExists($sPathOut) Then DirCreate($sPathOut)
						If $nBin = 2 Or $nBin = 3 Then                              ; $InEncoder = 'WEBP'
							$iProgFile += 10
							GUICtrlSetData($ProgFile, $iProgFile)
							$hImage = _DecodeFromWebP($pathWebP, $sPicsIn, $sPicsOut, $clsid)
;~ 							ConsoleWrite('2, 3' & @CRLF)
						EndIf
						If $nBin = 3 Then
							$iProgFile += 10
							GUICtrlSetData($ProgFile, $iProgFile)
							$hImage = _Resize(_IsChecked($Ratio), GUICtrlRead($Width), GUICtrlRead($Height), GUICtrlRead($pxpercent), 'none', $OutEncoder, $InEncoder, $hImage, $sPicsIn)
;~ 							ConsoleWrite('only 3' & @CRLF)
						EndIf
						If $nBin = 5 Then
							$iProgFile += 10
							GUICtrlSetData($ProgFile, $iProgFile)
							$hImage = _GDIPlus_ImageLoadFromFile($sPicsIn)
							$iProgFile += 10
							GUICtrlSetData($ProgFile, $iProgFile)
							$WidthHeight = _Resize(_IsChecked($Ratio), GUICtrlRead($Width), GUICtrlRead($Height), GUICtrlRead($pxpercent), 'none', $OutEncoder, $InEncoder, $hImage)
;~ 							ConsoleWrite('only5' & @CRLF)
						EndIf
						If $nBin = 7 Then
							$iProgFile += 10
							GUICtrlSetData($ProgFile, $iProgFile)
							$WidthHeight = _Resize(_IsChecked($Ratio), GUICtrlRead($Width), GUICtrlRead($Height), GUICtrlRead($pxpercent), 'none', $OutEncoder, $InEncoder, '', $sPicsIn)
;~ 							ConsoleWrite('only 7' & @CRLF)
						EndIf
						If $nBin = 2 Or $nBin = 4 Or $nBin = 6 Then      ; _IsChecked($Resizing) = False
							$iProgFile += 10
							GUICtrlSetData($ProgFile, $iProgFile)
							$WidthHeight[0] = ''
							$WidthHeight[1] = ''
;~ 							ConsoleWrite('2, 4, 6' & @CRLF)
						EndIf
						If $nBin >= 4 And $nBin <= 7 Then          ; $OutEncoder = 'WEBP'
							$iProgFile += 10
							GUICtrlSetData($ProgFile, $iProgFile)
							_EncodeToWebP($pathWebP, $sPicsIn, $sPicsOut, _IsChecked($Lossless), GUICtrlRead($JPGQlty), $WidthHeight[0], $WidthHeight[1])
;~ 							ConsoleWrite('4 To 7' & @CRLF)
						EndIf
						If $nBin = 1 Then      ;  $OutEncoder <> 'WEBP',   _IsChecked($Resizing)
							$iProgFile += 10
							GUICtrlSetData($ProgFile, $iProgFile)
							$hImage = _GDIPlus_ImageLoadFromFile($sPicsIn)
							$iProgFile += 10
							GUICtrlSetData($ProgFile, $iProgFile)
							$hImage = _Resize(_IsChecked($Ratio), GUICtrlRead($Width), GUICtrlRead($Height), GUICtrlRead($pxpercent), $iInterpolation, $OutEncoder, $InEncoder, $hImage, '')
;~ 							ConsoleWrite('1' & @CRLF)
						EndIf
						If $nBin = 0 Then     ;  $OutEncoder <> 'WEBP', $InEncoder <> 'WEBP',  _IsChecked($Resizing) = false
							$iProgFile += 10
							GUICtrlSetData($ProgFile, $iProgFile)
							$hImage = _GDIPlus_ImageLoadFromFile($sPicsIn)
;~ 							ConsoleWrite('0' & @CRLF)
						EndIf
						If $nBin >= 0 And $nBin <= 3 Then        ; $OutEncoder <> 'WEBP'
							$iProgFile += 10
							GUICtrlSetData($ProgFile, $iProgFile)
							_GDIPlus_ImageSaveToFileEx($hImage, $sPicsOut, $clsid, $Param)
							_GDIPlus_ImageDispose($hImage)
;~ 							ConsoleWrite('0 To 3' & @CRLF)
						EndIf
						$iProgFile += 10
						GUICtrlSetData($ProgFile, $iProgFile)
						FileSetTime($sPicsOut, FileGetTime($sPicsIn, 0, 1), 0)
						FileSetTime($sPicsOut, FileGetTime($sPicsIn, 1, 1), 1)
						GUICtrlSetData($ProgFile, 100)
					Next
					_GDIPlus_Shutdown()
					_ExtMsgBox(64, 0, "Done!", "Done!", 0, $Conv)
				EndIf
				GUISetState(@SW_HIDE, $Form1)
			EndIf
	EndSelect
WEnd

Func _checkValue($iValue)
	If $iValue > 100 Then $iValue = 100
	If $iValue < 1 Then $iValue = 1
	Return $iValue
EndFunc   ;==>_checkValue

Func _IsChecked($idControlID)
	Return BitAND(GUICtrlRead($idControlID), $GUI_CHECKED) = $GUI_CHECKED
EndFunc   ;==>_IsChecked

Func _CheckWebP()
	Local $Count = 0
	$TempDir = @LocalAppDataDir & '\temp\libwebp-1.6.0-windows-x64\bin'
	If Not (FileExists($TempDir & '\dwebp.exe')) Or Not (FileExists($TempDir & '\cwebp.exe')) Or Not (FileExists($TempDir & '\webpinfo.exe')) Then
		DirCreate(@LocalAppDataDir & '\temp\libwebp-1.6.0-windows-x64\bin')
		$Count += FileInstall('libwebp-1.6.0-windows-x64\bin\cwebp.exe', $TempDir & '\cwebp.exe', $FC_OVERWRITE)
		$Count += FileInstall('libwebp-1.6.0-windows-x64\bin\dwebp.exe', $TempDir & '\dwebp.exe', $FC_OVERWRITE)
		$Count += FileInstall('libwebp-1.6.0-windows-x64\bin\webpinfo.exe', $TempDir & '\webpinfo.exe', $FC_OVERWRITE)
	Else
		$Count = 3
	EndIf
	If $Count = 3 Then
;~ 		ConsoleWrite($TempDir & @CRLF)
		Return $TempDir
	Else
		Return ''
	EndIf
EndFunc   ;==>_CheckWebP

Func _FindPathName(ByRef $aRet, $sPath, $sFindFile, $bSubFolder = 0)
	Local $sSubFolderPath, $iIndex, $aFolders
	If Not IsArray($aRet) Then Return SetError(1, 0, -1)
	$aFile = _FileListToArray($sPath, $sFindFile, $FLTA_FILES, 1)
	If Not (@error) Then ; no files
		$aRet[0] = _ArrayConcatenate($aRet, $aFile, 1)
	EndIf
	GUICtrlSetData($Label2, 'Preparing files... ' & $aRet[0])
	$aFolders = _FileListToArray($sPath, "*", $FLTA_FOLDERS)
	If $bSubFolder Then
		If Not (@error) Then ; no folders
			For $i = 1 To $aFolders[0]
				$sSubFolderPath = $sPath & "\" & $aFolders[$i]
				$aRet[0] = _FindPathName($aRet, $sSubFolderPath, $sFindFile, $bSubFolder)
			Next
		EndIf
	EndIf
	$aRet[0] = UBound($aRet) - 1
	Return $aRet[0]
EndFunc   ;==>_FindPathName

Func _EncodeToWebP($sspathWebP, $ssPicsIn, $ssPicsOut, $bbLossless, $sQuality, $iWidth = '', $iHeight = '')
	GUICtrlSetData($ProgFile, 10)
	$ssspathWebP = $sspathWebP & '\cwebp.exe'
	$ssPicsIn = '"' & $ssPicsIn & '"'
	$ssPicsOut = '-o "' & $ssPicsOut & '"'
	$sParameter = '-mt -quiet -q ' & $sQuality
	If $bbLossless Then
		$sParameter &= ' -lossless'
	EndIf
	If $iWidth <> '' Or $iHeight <> '' Then
		$sParameter &= ' -resize ' & $iWidth & ' ' & $iHeight
	EndIf
	$cmd = $ssspathWebP & ' ' & $sParameter & ' ' & $ssPicsIn & ' ' & $ssPicsOut
;~ 	ConsoleWrite($cmd & @CRLF)
	GUICtrlSetData($ProgFile, 40)
	RunWait(@ComSpec & " /c " & $cmd, @SystemDir, @SW_HIDE)
	GUICtrlSetData($ProgFile, 60)
EndFunc   ;==>_EncodeToWebP

Func _DecodeFromWebP($spathWebP, $ssPicsIn, $ssPicsOut, $cclsid)
	Local $sOutput = '', $sParam
	$ssspathWebP = $spathWebP & '\dwebp.exe'
	$sParameter = '-mt -quiet'
	$cmd = $ssspathWebP & ' ' & $sParameter & ' ' & $ssPicsIn & ' -o -'
	Local $iPID = Run(@ComSpec & " /c " & $cmd, @SystemDir, @SW_HIDE, $STDERR_MERGED)
	While 1
		$sOutput &= StdoutRead($iPID)
		If @error Then ; Exit the loop if the process closes or StdoutRead returns an error.
			ExitLoop
		EndIf
	WEnd
	$sOutput = StringToBinary($sOutput)     ; Convert the string to binary.
	$hGdi = _GDIPlus_BitmapCreateFromMemory($sOutput)
	Return $hGdi
;~ 	_GDIPlus_ImageSaveToFileEx($hGdi, $ssPicsOut, $cclsid)
;~ 	_GDIPlus_BitmapDispose($hGdi)
	#cs
	    Ok = 0,
	    GenericError = 1,
	    InvalidParameter = 2,
	    OutOfMemory = 3,
	    ObjectBusy = 4,
	    InsufficientBuffer = 5,
	    NotImplemented = 6,
	    Win32Error = 7,
	    WrongState = 8,
	    Aborted = 9,
	    FileNotFound = 10,
	    ValueOverflow = 11,
	    AccessDenied = 12,
	    UnknownImageFormat = 13,
	    FontFamilyNotFound = 14,
	    FontStyleNotFound = 15,
	    NotTrueTypeFont = 16,
	    UnsupportedGdiplusVersion = 17,
	    GdiplusNotInitialized = 18,
	    PropertyNotFound = 19,
	    PropertyNotSupported = 20,
	#ce
EndFunc   ;==>_DecodeFromWebP

;================================== FUNC _Resize ===========================================
; Resize the image: calculate the Width and Height according of the inputs
;
; Output: if it is a WEBP: Width and Height in $aDim (Width = $aDim[0], Height = $aDim[1])
;         Otherwise return $hhImage from _GDIPlus_ImageResize
;
;===========================================================================================
Func _Resize($bRatio, $iiWidth, $iiHeight, $iipxpercent, $iInterpolation, $ssOutEncoder, $ssInEncoder, $hhImage = '', $ssPicsIn = '')
	Local $aDim[2]
	If $ssInEncoder = 'WEBP' Then
		$sOutput = ''
		Local $iPID = Run($pathWebP & '\webpinfo.exe ' & $ssPicsIn, @SystemDir, @SW_HIDE, $STDERR_MERGED)
		While 1
			$sOutput &= StdoutRead($iPID)
			If @error Then ; Exit the loop if the process closes or StdoutRead returns an error.
				ExitLoop
			EndIf
		WEnd
		$aArray = StringSplit($sOutput, @CRLF)
		For $i = 1 To UBound($aArray) - 1
			If StringInStr($aArray[$i], 'Width') Then
				$aSplit = StringSplit(StringStripWS($aArray[$i], 8), ':')
				$iWidthIm = $aSplit[2]
			ElseIf StringInStr($aArray[$i], 'Height') Then
				$aSplit = StringSplit(StringStripWS($aArray[$i], 8), ':')
				$iHeightIm = $aSplit[2]
			EndIf
		Next
	Else
		$aDim = _GDIPlus_ImageGetDimension($hhImage)
		$iWidthIm = $aDim[0]
		$iHeightIm = $aDim[1]
	EndIf
	If $bRatio Then  ;ratio checked
		$fRatio = $iWidthIm / $iHeightIm  ; width / height
		If $iipxpercent = 'px' Then  ; px with ratio
			If $iiWidth = '' Then
				$iWidth = Round($iiHeight * $fRatio)
				$iHeight = $iiHeight
			Else
				$iWidth = $iiWidth
				$iHeight = Round($iiWidth / $fRatio)
			EndIf
		EndIf
	Else ;ratio Not checked
		If $iipxpercent = 'px' Then ; px in no ratio
			$iWidth = $iiWidth
			$iHeight = $iiHeight
		EndIf
	EndIf
	If $iipxpercent = '%' Then ; pourcent
		$iHeight = Round($iHeightIm * ($iiHeight / 100))
		$iWidth = Round($iWidthIm * ($iiWidth / 100))
	EndIf
	If $ssOutEncoder = 'WEBP' Then
		$aDim[0] = $iWidth
		$aDim[1] = $iHeight
;~ 		ConsoleWrite('Return $aDim' & @CRLF)
		Return $aDim
	Else
		$hhImage = _GDIPlus_ImageResize($hhImage, $iWidth, $iHeight, $iInterpolation) ;resized image
;~ 		ConsoleWrite('Return $hhimage' & @CRLF)
		Return $hhImage
	EndIf
EndFunc   ;==>_Resize

Func _CheckResize($bChecked)
	If $bChecked Then
		GUICtrlSetState($Ratio, $GUI_ENABLE)
		GUICtrlSetState($Width, $GUI_ENABLE)
		GUICtrlSetState($Label4, $GUI_ENABLE)
		GUICtrlSetState($Height, $GUI_ENABLE)
		GUICtrlSetState($Label5, $GUI_ENABLE)
		GUICtrlSetState($Label3, $GUI_ENABLE)
		GUICtrlSetState($Interpolation, $GUI_ENABLE)
		GUICtrlSetState($pxpercent, $GUI_ENABLE)
		GUICtrlSetState($Label6, $GUI_ENABLE)
		GUICtrlSetState($Label7, $GUI_ENABLE)
		If $OutEncoder = 'WEBP' Then
			GUICtrlSetState($Label3, $GUI_DISABLE)
			GUICtrlSetState($Interpolation, $GUI_DISABLE)
		Else
			GUICtrlSetState($Label3, $GUI_ENABLE)
			GUICtrlSetState($Interpolation, $GUI_ENABLE)
		EndIf
	Else
		GUICtrlSetState($Ratio, $GUI_DISABLE)
		GUICtrlSetState($Width, $GUI_DISABLE)
		GUICtrlSetState($Label4, $GUI_DISABLE)
		GUICtrlSetState($Height, $GUI_DISABLE)
		GUICtrlSetState($Label5, $GUI_DISABLE)
		GUICtrlSetState($Label3, $GUI_DISABLE)
		GUICtrlSetState($Label6, $GUI_DISABLE)
		GUICtrlSetState($Label7, $GUI_DISABLE)
		GUICtrlSetState($Interpolation, $GUI_DISABLE)
		GUICtrlSetState($pxpercent, $GUI_DISABLE)
	EndIf
EndFunc   ;==>_CheckResize

Func _FindExtention($_sPath, $_sDecoder)
	Dim $_aArray[1]
	Local $sDrive = '', $sDir = '', $sFileName = '', $sExtension = ''
	$_aDecoder = StringSplit($_sDecoder, '|')
	_FindPathName($_aArray, $_sPath, '*', 0)
	For $i = 1 To UBound($_aArray) - 1
		$_aPath = _PathSplit($_aArray[$i], $sDrive, $sDir, $sFileName, $sExtension)
		$extension = StringTrimLeft($_aPath[$PATH_EXTENSION], 1)
		For $j = 1 To UBound($_aDecoder) - 1
			If $extension = $_aDecoder[$j] Then
				Return $_aDecoder[$j]
			EndIf
		Next
	Next
	Return ''
EndFunc   ;==>_FindExtention
