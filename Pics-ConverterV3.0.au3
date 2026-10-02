#Region ;**** Directives created by AutoIt3Wrapper_GUI ****
#AutoIt3Wrapper_Icon=..\AutoItv11.ico
#AutoIt3Wrapper_Res_Comment=Convert and resize from/to *NEW* WEBP, JPG, BMP, GIF, PNG,...
#AutoIt3Wrapper_Res_Description=Convert and resize from/to *NEW* WEBP, JPG, BMP, GIF, PNG,...
#AutoIt3Wrapper_Res_Fileversion=3.0.3.0
#AutoIt3Wrapper_Res_ProductName=Pics Converter V3
#AutoIt3Wrapper_Res_ProductVersion=3.0.3.0
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
#    Last compile at : 2026/10/02 13:29:03
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

	Limitation: pictures cannot be at the root of f.ex D:. Put it in a folder!

	Bug: not known

	To Do:	how to keep metadata? WEBP: " -metadata all" a voir dans la doc

	V3.0.3.0	02.10.2026:
				Big optimisation, improvement whith Claude AI and many bugs resolving
	V3.0.2.1	04.12.2025:
				Slight improve
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

$sVersion = 'V3.0.3.0'
$head = 'Pics Conversion ' & $sVersion

Local $Param = 0, $Decoder, $ToCombo, $ToComboOut, $OldOutEncoder, $Oldpxpercent, $bGo = True, $EncoderExt[1] = [0], $DecoderExt[1] = [0]
Local $OldValSlider = 0, $OldJPGQuality = 100, $OldHeight, $Oldwidth, $OldCheckRatio, $OldLossless, $OldResize, $Parameter, $WidthHeight[2]
Dim $aInterpolation[2][7] = [[$GDIP_INTERPOLATIONMODE_HIGHQUALITYBICUBIC, $GDIP_INTERPOLATIONMODE_HIGHQUALITYBILINEAR, $GDIP_INTERPOLATIONMODE_NEARESTNEIGHBOR, $GDIP_INTERPOLATIONMODE_BICUBIC, $GDIP_INTERPOLATIONMODE_BILINEAR, $GDIP_INTERPOLATIONMODE_HIGHQUALITY, $GDIP_INTERPOLATIONMODE_LOWQUALITY], ['Bicubic HQ (default)', 'Bilinear HQ', 'Nearest neighbor', 'Bicubic (low)', 'Bilinear (low)', 'High-quality', 'Low-quality']]
Global $pathWebP = _CheckWebP(), $bFixQuality = False, $bFixWidth = False, $bFixHeight = False, $Label2Form1, $bIsSaved, $iSaved = 0

GUIRegisterMsg($WM_COMMAND, "_WM_COMMAND")

_GDIPlus_Startup()
$Decoder = _GDIPlus_Decoders()
$Encoder = _GDIPlus_Encoders()
_GDIPlus_Shutdown()

If $pathWebP <> '' Then
	$Encoder[0][0] = _ArrayAdd($Encoder, '*.WEBP', 6)
	$Decoder[0][0] = _ArrayAdd($Decoder, '*.WEBP', 6)
EndIf

For $i = 1 To $Encoder[0][0]
	$Split = StringSplit($Encoder[$i][6], ";")
	For $j = 1 To $Split[0]
		$ToComboOut &= StringTrimLeft($Split[$j], 2) & "|"
	Next
Next
$ToComboOut = StringTrimRight($ToComboOut, 1)
$EncoderExt = _ArrayFromString($ToComboOut)
_ArraySort($EncoderExt)
$ToComboOut = _ArrayToString($EncoderExt)

For $i = 1 To $Decoder[0][0]
	$Split = StringSplit($Decoder[$i][6], ";")
	For $j = 1 To $Split[0]
		$ToCombo &= StringTrimLeft($Split[$j], 2) & "|"
	Next
Next
$ToCombo = StringTrimRight($ToCombo, 1)
$DecoderExt = _ArrayFromString($ToCombo)
_ArraySort($DecoderExt)
$ToCombo = _ArrayToString($DecoderExt)

$Conv = GUICreate($head, 570, 210, -1, -1)
$Group1 = GUICtrlCreateGroup(" Input ", 5, 5, 140, 145)
$InputEncoder = GUICtrlCreateCombo("", 15, 120, 120, 25, BitOR($CBS_DROPDOWNLIST, $CBS_AUTOHSCROLL))
GUICtrlSetData(-1, $ToCombo)
$InputFolder = GUICtrlCreateInput("Input Folder", 15, 25, 120, 21)
$BrowseInput = GUICtrlCreateButton("Browse...", 60, 50, 75, 25, $WS_GROUP)
$Subfolder = GUICtrlCreateCheckbox("Subfolers ?", 60, 77, 75, 21)
$Label9 = GUICtrlCreateLabel("Convert from:", 15, 100, 67, 17)
GUICtrlCreateGroup("", -99, -99, 1, 1)

$Group4 = GUICtrlCreateGroup(" Resize ", 155, 5, 160, 145)
$Resizing = GUICtrlCreateCheckbox("Resizing", 165, 25, 70, 18)
$pxpercent = GUICtrlCreateCombo('px', 250, 25, 40, 25, BitOR($CBS_DROPDOWNLIST, $CBS_AUTOHSCROLL))
GUICtrlSetData(-1, '%')
$Ratio = GUICtrlCreateCheckbox("Keep aspect ratio", 175, 47, 115, 20)
GUICtrlSetState(-1, $GUI_CHECKED)

$Width = GUICtrlCreateInput("", 170, 72, 49, 21, $ES_NUMBER)
$Label4 = GUICtrlCreateLabel("px", 220, 77, 15, 17)
$Height = GUICtrlCreateInput("", 248, 72, 49, 21, $ES_NUMBER)
$Label5 = GUICtrlCreateLabel("px", 298, 77, 15, 17)
$Label6 = GUICtrlCreateLabel("w:", 158, 77, 10, 17)
$Label7 = GUICtrlCreateLabel("h:", 238, 77, 10, 17)

$Label3 = GUICtrlCreateLabel("Interpolation mode:", 168, 100, 104, 17)
$Interpolation = GUICtrlCreateCombo($aInterpolation[1][0], 165, 120, 120, 25, BitOR($CBS_DROPDOWNLIST, $CBS_AUTOHSCROLL))
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
$OutputEncoder = GUICtrlCreateCombo("", 335, 120, 120, 25, BitOR($CBS_DROPDOWNLIST, $CBS_AUTOHSCROLL))
GUICtrlSetData(-1, $ToComboOut)
$OutputFolder = GUICtrlCreateInput("Output Folder", 335, 25, 120, 21)
$BrowseOutput = GUICtrlCreateButton("Browse...", 380, 50, 75, 25, $WS_GROUP)
$Label1 = GUICtrlCreateLabel("Convert to:", 335, 100, 56, 17)
GUICtrlCreateGroup("", -99, -99, 1, 1)

$Group3 = GUICtrlCreateGroup(" Quality ", 475, 5, 90, 145)
$Lossless = GUICtrlCreateCheckbox('Lossless', 480, 25, 60, 25)
$Slider = GUICtrlCreateSlider(515, 47, 35, 100, BitOR($TBS_VERT, $TBS_TOP, $TBS_LEFT))
$JPGQlty = GUICtrlCreateInput("100", 485, 87, 30, 21, $ES_NUMBER)

GUICtrlSetState($Group3, $GUI_ENABLE)
GUICtrlSetState($Slider, $GUI_DISABLE)
GUICtrlSetState($JPGQlty, $GUI_DISABLE)
GUICtrlSetState($Lossless, $GUI_DISABLE)

GUICtrlCreateGroup("", -99, -99, 1, 1)

$GO = GUICtrlCreateButton("Convert", 185, 160, 200, 40, $WS_GROUP)
GUISetState(@SW_SHOW)

While 1
	$nMsg = GUIGetMsg()

	If $bFixQuality Then
		$bFixQuality = False
		$iQ = GUICtrlRead($JPGQlty)
		If $iQ = '' Then $iQ = 100
		$iQ = _checkValue($iQ)
		GUICtrlSetData($JPGQlty, $iQ)
		GUICtrlSetData($Slider, 100 - $iQ)
		$OldJPGQuality = $iQ
	EndIf
	If $bFixWidth Then
		$bFixWidth = False
		If GUICtrlRead($pxpercent) = '%' Then
			$iV = _FixPercent(GUICtrlRead($Width))
			GUICtrlSetData($Width, $iV)
			If _IsChecked($Ratio) Then GUICtrlSetData($Height, $iV)
			$Oldwidth = GUICtrlRead($Width)
			$OldHeight = GUICtrlRead($Height)
		EndIf
	EndIf
	If $bFixHeight Then
		$bFixHeight = False
		If GUICtrlRead($pxpercent) = '%' Then
			$iV = _FixPercent(GUICtrlRead($Height))
			GUICtrlSetData($Height, $iV)
			If _IsChecked($Ratio) Then GUICtrlSetData($Width, $iV)
			$Oldwidth = GUICtrlRead($Width)
			$OldHeight = GUICtrlRead($Height)
		EndIf
	EndIf
	$OutEncoder = GUICtrlRead($OutputEncoder)
	$sPxpercent = GUICtrlRead($pxpercent)
	$ValSlider = GUICtrlRead($Slider)
	$JPGQuality = GUICtrlRead($JPGQlty)
	$iHeight = GUICtrlRead($Height)
	$iWidth = GUICtrlRead($Width)
	;---------------- SELECT ----------------------------
	Select
		Case $nMsg = $GUI_EVENT_CLOSE
			Exit
		Case $nMsg = $BrowseInput
			$sInFold = GUICtrlRead($InputFolder)
			If $sInFold = "Input Folder" Then $sInFold = ""
			$sInFold = FileSelectFolder("Choose a folder", $sInFold, 7, '', $Conv)
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
			If $sOutFold = "Output Folder" Then $sOutFold = ""
			$sOutFold = FileSelectFolder("Choose a folder", $sOutFold, 1, '', $Conv)
			If $sOutFold <> "" Then
				$OutFold = $sOutFold
				GUICtrlSetData($OutputFolder, $OutFold)
			EndIf
		Case $OutEncoder <> $OldOutEncoder
			If (Not _IsJpeg($OutEncoder)) And ($OutEncoder <> 'WEBP') Then
				GUICtrlSetState($Lossless, $GUI_UNCHECKED)
				GUICtrlSetState($Lossless, $GUI_DISABLE)
				GUICtrlSetState($Group3, $GUI_DISABLE)
				GUICtrlSetState($Slider, $GUI_DISABLE)
				GUICtrlSetState($JPGQlty, $GUI_DISABLE)
				GUICtrlSetState($Lossless, $GUI_DISABLE)
				GUICtrlSetState($Resizing, $GUI_ENABLE)
			Else
				If _IsJpeg($OutEncoder) Then
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
			If _IsChecked($Resizing) Then
				_CheckResize(True)
			EndIf
			$OldOutEncoder = $OutEncoder
		Case $ValSlider <> $OldValSlider
			GUICtrlSetData($JPGQlty, 100 - $ValSlider)
			$OldValSlider = $ValSlider
		Case $JPGQuality <> $OldJPGQuality
			If $JPGQuality <> '' And $JPGQuality >= 1 And $JPGQuality <= 100 Then
				GUICtrlSetData($Slider, 100 - $JPGQuality)
			EndIf
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
				If $iWidth <> '' Then GUICtrlSetData($Width, _checkValue($iWidth))
				If $iHeight <> '' Then GUICtrlSetData($Height, _checkValue($iHeight))
				If _IsChecked($Ratio) And _ValidPercent(GUICtrlRead($Width)) Then GUICtrlSetData($Height, GUICtrlRead($Width))
				$Oldwidth = GUICtrlRead($Width)
				$OldHeight = GUICtrlRead($Height)
			EndIf
			$Oldpxpercent = $sPxpercent
		Case ($iWidth <> $Oldwidth Or $iHeight <> $OldHeight) And GUICtrlRead($pxpercent) = '%'
			If _IsChecked($Ratio) Then
				If $iWidth <> $Oldwidth Then
					If _ValidPercent($iWidth) Then GUICtrlSetData($Height, $iWidth)
				Else
					If _ValidPercent($iHeight) Then GUICtrlSetData($Width, $iHeight)
				EndIf
			EndIf
			$Oldwidth = GUICtrlRead($Width)
			$OldHeight = GUICtrlRead($Height)
		Case _IsChecked($Ratio) <> $OldCheckRatio And GUICtrlRead($pxpercent) = '%'
			$OldCheckRatio = _IsChecked($Ratio)
			If $OldCheckRatio And _ValidPercent($iWidth) Then GUICtrlSetData($Height, $iWidth)
			$Oldwidth = GUICtrlRead($Width)
			$OldHeight = GUICtrlRead($Height)
			;	------------------------------------- Convert ----------------------------------------
		Case $nMsg = $GO
			$bGo = True
			$Param = 0
			$InPath = GUICtrlRead($InputFolder)
			$OutPath = GUICtrlRead($OutputFolder)
			$InEncoder = GUICtrlRead($InputEncoder)
			$OutEncoder = GUICtrlRead($OutputEncoder)
			If Not (StringInStr($InPath, "\")) Or Not (StringInStr($OutPath, "\")) Then
				_ExtMsgBox(16, 0, "Caution!", "Please select a folder!", 0, $Conv)
				$bGo = False
			ElseIf $InEncoder = '' Or $OutEncoder = '' Then
				_ExtMsgBox(16, 0, "Caution!", "Input and/or Output Encoder cannot be empty", 0, $Conv)
				$bGo = False
			ElseIf GUICtrlRead($Width) = '' And GUICtrlRead($Height) = '' And _IsChecked($Resizing) Then
				_ExtMsgBox(16, 0, "Caution!", "Width and/or height cannot be empty", 0, $Conv)
				$bGo = False
			ElseIf $InPath = $OutPath And $InEncoder = $OutEncoder And _IsChecked($Resizing) Then
				If _ExtMsgBox(48, 4, "Warning!", "The input and output folder are the same!" & @CRLF & _
						"Continuing will overwrite input pictures!" & @CRLF & "Do you want to continue?", 0, $Conv) <> 1 Then $bGo = False
			ElseIf $InPath = $OutPath And $InEncoder = $OutEncoder And Not (_IsChecked($Resizing)) Then
				_ExtMsgBox(16, 0, "Caution!", "The input and output folder are the same!" & @CRLF & _
						"Please choose a different folder or a different encoder", 0, $Conv)
				$bGo = False
			ElseIf _IsChecked($Resizing) And Not _IsChecked($Ratio) And (GUICtrlRead($Width) = '' Or GUICtrlRead($Height) = '') Then
				_ExtMsgBox(16, 0, "Caution!", "Width and height must both be filled", 0, $Conv)
				$bGo = False
			EndIf
			If $bGo Then
;~ 				do the conversion process...
;~ 				do the progress bar GUI
				$aPos = WinGetPos($Conv)
				$iWinWidth = 550
				$iWinHeight = 135
				$Form1 = GUICreate("", $iWinWidth, $iWinHeight, ($aPos[0] + ($aPos[2] / 2)) - ($iWinWidth / 2), ($aPos[1] + ($aPos[3] / 2)) - ($iWinHeight / 2), BitOR($WS_POPUP, $WS_BORDER), $WS_EX_TOOLWINDOW, $Conv)
				Global $ProgFile = GUICtrlCreateProgress(10, 10, $iWinWidth - 20, 20, $PBS_SMOOTH)
				$Label2Form1 = GUICtrlCreateLabel("", 10, 35, $iWinWidth - 20, 20, $SS_LEFT)
				$ProgAll = GUICtrlCreateProgress(10, 60, $iWinWidth - 20, 20, $PBS_SMOOTH)
				$Label3Form1 = GUICtrlCreateLabel("", 10, 85, $iWinWidth - 20, 20, $SS_LEFT)
				$Label1Form1 = GUICtrlCreateLabel("", 10, 110, $iWinWidth - 20, 20, $SS_LEFT)
				GUISetState(@SW_SHOW)
				If _IsJpeg($OutEncoder) Then ; Set JPG quality
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
					_ExtMsgBox(16, 0, "Caution!", "No files found or invalid path or wrong input format!", 0, $Conv)
				Else
					_GDIPlus_Startup()
					If $OutEncoder <> 'WEBP' Then
						$clsid = _GDIPlus_EncodersGetCLSID($OutEncoder)
					EndIf
					$nBin = 0
					If $OutEncoder = 'WEBP' Then $nBin += 4  ; cwebp : compresse un fichier image en fichier WebP
					If $InEncoder = 'WEBP' Then $nBin += 2  ; dwebp : décompresser un fichier WebP dans un fichier image
					If _IsChecked($Resizing) Then $nBin += 1 ; Resizing = True
;~ 					ConsoleWrite($nBin & @CRLF)
					For $i = 1 To $FileList[0]
						$iProgFile = 0
						GUICtrlSetData($ProgFile, $iProgFile)
						GUICtrlSetData($ProgAll, ($i / $FileList[0]) * 100)
						$sPicsIn = $FileList[$i]
						$sRel = StringTrimLeft($sPicsIn, StringLen($InPath))          ; partie relative
						$sRel = StringRegExpReplace($sRel, '(?i)\.' & $InEncoder & '$', '.' & $OutEncoder) ; nouvelle extension
						$sPicsOut = $OutPath & $sRel                                  ; fichier de sortie
						$sPathOut = StringLeft($sPicsOut, StringInStr($sPicsOut, '\', 0, -1) - 1)
						GUICtrlSetData($Label2Form1, $sPicsIn)
						GUICtrlSetData($Label3Form1, $sPicsOut)
						GUICtrlSetData($Label1Form1, $i & " / " & $FileList[0] & ' - ' & Round(($i / $FileList[0]) * 100, 1) & '%')
;~ 						$sPathOut = StringReplace($sPicsOut, $aPath[UBound($aPath) - 1], '')
						If Not FileExists($sPathOut) Then DirCreate($sPathOut)
						If $nBin = 2 Or $nBin = 3 Then                                                ; $InEncoder = 'WEBP'
							$iProgFile += 10
							GUICtrlSetData($ProgFile, $iProgFile)
							$hImage = _DecodeFromWebP($pathWebP, $sPicsIn, $clsid)
							If Not $hImage Then
								$iSaved += 1
								ContinueLoop
							EndIf
;~ 							ConsoleWrite('2, 3' & @CRLF)
						EndIf
						If $nBin = 3 Then                                                            ; $InEncoder = 'WEBP' and Resize = true (1)
							$iProgFile += 10
							GUICtrlSetData($ProgFile, $iProgFile)
							$hOld = $hImage
							$hImage = _Resize(_IsChecked($Ratio), GUICtrlRead($Width), GUICtrlRead($Height), GUICtrlRead($pxpercent), $iInterpolation, $OutEncoder, $InEncoder, $hImage, $sPicsIn)
							_GDIPlus_ImageDispose($hOld)
							If Not $hImage Then
								$iSaved += 1
								ContinueLoop
							EndIf
;~ 							ConsoleWrite('only 3' & @CRLF)
						EndIf
						If $nBin = 5 Then                                                            ; $OutEncoder = 'WEBP' and Resize = true (1)
							$iProgFile += 10
							GUICtrlSetData($ProgFile, $iProgFile)
							$hImage = _GDIPlus_ImageLoadFromFile($sPicsIn)
							$iProgFile += 10
							GUICtrlSetData($ProgFile, $iProgFile)
							If Not $hImage Then
								$iSaved += 1
								ContinueLoop
							EndIf
							$WidthHeight = _Resize(_IsChecked($Ratio), GUICtrlRead($Width), GUICtrlRead($Height), GUICtrlRead($pxpercent), 'none', $OutEncoder, $InEncoder, $hImage)
							_GDIPlus_ImageDispose($hImage)
;~ 							ConsoleWrite('only5' & @CRLF)
						EndIf
						If $nBin = 7 Then                                                            ; $OutEncoder = 'WEBP' and $InEncoder = 'WEBP' andResize = true (1)
							$iProgFile += 10
							GUICtrlSetData($ProgFile, $iProgFile)
							$WidthHeight = _Resize(_IsChecked($Ratio), GUICtrlRead($Width), GUICtrlRead($Height), GUICtrlRead($pxpercent), 'none', $OutEncoder, $InEncoder, '', $sPicsIn)
;~ 							ConsoleWrite('only 7' & @CRLF)
						EndIf
						If $nBin = 2 Or $nBin = 4 Or $nBin = 6 Then                                    ; _IsChecked($Resizing) = False (0)
							$iProgFile += 10
							GUICtrlSetData($ProgFile, $iProgFile)
							$WidthHeight[0] = ''
							$WidthHeight[1] = ''
;~ 							ConsoleWrite('2, 4, 6' & @CRLF)
						EndIf
						If $nBin >= 4 And $nBin <= 7 Then                                            ; $OutEncoder = 'WEBP'
							$iProgFile += 10
							GUICtrlSetData($ProgFile, $iProgFile)
							$bIsSaved = _EncodeToWebP($pathWebP, $sPicsIn, $sPicsOut, _IsChecked($Lossless), GUICtrlRead($JPGQlty), $WidthHeight[0], $WidthHeight[1])
							If Not $bIsSaved Then $iSaved += 1
;~ 							ConsoleWrite('4 To 7' & @CRLF)
						EndIf
						If $nBin = 1 Then                                                            ;  $OutEncoder <> 'WEBP', _IsChecked($Resizing)
							$iProgFile += 10
							GUICtrlSetData($ProgFile, $iProgFile)
							$hImage = _GDIPlus_ImageLoadFromFile($sPicsIn)
							If Not $hImage Then
								$iSaved += 1
								ContinueLoop
							EndIf
							$hOld = $hImage
							$iProgFile += 10
							GUICtrlSetData($ProgFile, $iProgFile)
							$hImage = _Resize(_IsChecked($Ratio), GUICtrlRead($Width), GUICtrlRead($Height), GUICtrlRead($pxpercent), $iInterpolation, $OutEncoder, $InEncoder, $hImage, '')
							_GDIPlus_ImageDispose($hOld) ; On dispose l'image original (pas le resize)
							If Not $hImage Then
								$iSaved += 1
								ContinueLoop
							EndIf
;~ 							ConsoleWrite('1' & @CRLF)
						EndIf
						If $nBin = 0 Then                                                            ;  $OutEncoder <> 'WEBP', $InEncoder <> 'WEBP',  _IsChecked($Resizing) = false
							$iProgFile += 10
							GUICtrlSetData($ProgFile, $iProgFile)
							$hImage = _GDIPlus_ImageLoadFromFile($sPicsIn)
;~ 							ConsoleWrite('0' & @CRLF)
						EndIf
						If $nBin >= 0 And $nBin <= 3 Then                                            ; $OutEncoder <> 'WEBP'
							$iProgFile += 10
							GUICtrlSetData($ProgFile, $iProgFile)
							$bIsSaved = _GDIPlus_ImageSaveToFileEx($hImage, $sPicsOut, $clsid, $Param)
							If Not $bIsSaved Then $iSaved += 1
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
					If $iSaved > 0 Then
						_ExtMsgBox(64, 0, "Done!", "Done with " & $iSaved & " errors", 0, $Conv)
					Else
						_ExtMsgBox(64, 0, "Done!", "Done!", 0, $Conv)
					EndIf
					$iSaved = 0
				EndIf
				GUIDelete($Form1)
				$Label2Form1 = 0
			EndIf
	EndSelect
WEnd

; ================================================================= FUNCTIONS ================================================================

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

Func _EncodeToWebP($sspathWebP, $ssPicsIn, $ssPicsOut, $bbLossless, $sQuality, $iWidth = '', $iHeight = '')
	GUICtrlSetData($ProgFile, 10)
	$ssspathWebP = $sspathWebP & '\cwebp.exe'
	$ssPicsIn = '"' & $ssPicsIn & '"'
	$ssPicsOut = '-o "' & $ssPicsOut & '"'
	$sssspathWebP = '"' & $ssspathWebP & '"'
	$sParameter = '-mt -quiet -q ' & $sQuality
	If $bbLossless Then
		$sParameter &= ' -lossless'
	EndIf
	If $iWidth <> '' Or $iHeight <> '' Then
		$sParameter &= ' -resize ' & $iWidth & ' ' & $iHeight
	EndIf
	$cmd = $sssspathWebP & ' ' & $sParameter & ' ' & $ssPicsIn & ' ' & $ssPicsOut
;~ 	ConsoleWrite($cmd & @CRLF)
	GUICtrlSetData($ProgFile, 40)
	$ReturnCode = RunWait($cmd, @SystemDir, @SW_HIDE) ; quand cmd ( @ComSpec & " /c "&..) est lancé il supprime le 1er est le dernier guillemets
	GUICtrlSetData($ProgFile, 60)
	Return Not ($ReturnCode)
EndFunc   ;==>_EncodeToWebP

Func _DecodeFromWebP($spathWebP, $ssPicsIn, $cclsid)
	Local $sOutput = '', $sParam
	$ssspathWebP = $spathWebP & '\dwebp.exe'
	$ssPicsIn = '"' & $ssPicsIn & '"'
	$sssspathWebP = '"' & $ssspathWebP & '"'
	$sParameter = '-mt -quiet'
	$cmd = $sssspathWebP & ' ' & $sParameter & ' ' & $ssPicsIn & ' -o -'
;~ 	ConsoleWrite($cmd & @CRLF)
	Local $iPID = Run($cmd, @SystemDir, @SW_HIDE, $STDOUT_CHILD) ; quand cmd ( @ComSpec & " /c "&..) est lancé il supprime le 1er est le dernier guillemets
	While 1
		$sOutput &= StdoutRead($iPID)
		If @error Then ExitLoop ; Exit the loop if the process closes or StdoutRead returns an error.
		Sleep(10)
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
	If $ssInEncoder = 'WEBP' And $ssOutEncoder = 'WEBP' Then ; get dimentions
		$sOutput = ''
		Local $iPID = Run('"' & $pathWebP & '\webpinfo.exe' & '" "' & $ssPicsIn & '"', @SystemDir, @SW_HIDE, $STDERR_MERGED)
;~ 		ConsoleWrite('"' & $pathWebP & '\webpinfo.exe' & '" "' & $ssPicsIn & '"' & @CRLF)
		While 1
			$sOutput &= StdoutRead($iPID)
			If @error Then ExitLoop ; Exit the loop if the process closes or StdoutRead returns an error.
			Sleep(10)
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
	If $iWidthIm < 1 Or $iHeightIm < 1 Then
		_ExtMsgBox(16, 0, "Caution!", " The Dimensions cannot be determined. The image could be corrupted! The program will end", 0, $Conv)
		Exit
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
;~ 		Else ; code inutile
;~ 			GUICtrlSetState($Label3, $GUI_ENABLE)
;~ 			GUICtrlSetState($Interpolation, $GUI_ENABLE)
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
	Return $bChecked
EndFunc   ;==>_CheckResize

Func _FindExtention($_sPath, $_sDecoder)
	Local $aDec = StringSplit($_sDecoder, '|'), $sName, $iPos, $sExt
	Local $hSearch = FileFindFirstFile($_sPath & '\*.*')
	If $hSearch = -1 Then Return ''
	While 1
		$sName = FileFindNextFile($hSearch)
		If @error Then ExitLoop
		If @extended Then ContinueLoop                ; dossier
		$iPos = StringInStr($sName, '.', 0, -1)
		If $iPos = 0 Then ContinueLoop
		$sExt = StringTrimLeft($sName, $iPos)
		For $j = 1 To $aDec[0]
			If $aDec[$j] <> '' And $sExt = $aDec[$j] Then
				FileClose($hSearch)
				Return $aDec[$j]
			EndIf
		Next
	WEnd
	FileClose($hSearch)
	Return ''
EndFunc   ;==>_FindExtention

Func _FindPathName(ByRef $aRet, $sPath, $sFindFile, $bSubFolder = 0)
	Local $sSubFolderPath, $iIndex, $aFolders
	If Not IsArray($aRet) Then Return SetError(1, 0, -1)
	$aFile = _FileListToArray($sPath, $sFindFile, $FLTA_FILES, 1)
	If Not (@error) Then ; no files
		$aRet[0] = _ArrayConcatenate($aRet, $aFile, 1)
	EndIf
	GUICtrlSetData($Label2Form1, 'Preparing files... ' & $aRet[0])
	If $bSubFolder Then
		$aFolders = _FileListToArray($sPath, "*", $FLTA_FOLDERS)
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

Func _WM_COMMAND($hWnd, $iMsg, $wParam, $lParam)
	Local $iID = BitAND($wParam, 0xFFFF)         ; ID du contrôle
	Local $iCode = BitShift($wParam, 16)         ; code de notification
	If $iCode = $EN_KILLFOCUS Then
		Switch $iID
			Case $JPGQlty
				$bFixQuality = True
			Case $Width
				$bFixWidth = True
			Case $Height
				$bFixHeight = True
		EndSwitch
	EndIf
	Return $GUI_RUNDEFMSG
EndFunc   ;==>_WM_COMMAND

; Valeur valide en % ? (non vide, entre 1 et 100)
Func _ValidPercent($sVal)
	If $sVal = '' Then Return False
	Return Number($sVal) >= 1 And Number($sVal) <= 100
EndFunc   ;==>_ValidPercent

; Valeur finale à la perte du focus : vide -> 100, sinon bornée à 1-100
Func _FixPercent($sVal)
	If $sVal = '' Then Return 100
	Return _checkValue(Number($sVal))
EndFunc   ;==>_FixPercent

Func _IsJpeg($s)
	Return StringInStr('|JPG|JPEG|JPE|JFIF|', '|' & $s & '|') > 0
EndFunc   ;==>_IsJpeg
