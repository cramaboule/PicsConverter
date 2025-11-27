#Region ;**** Directives created by AutoIt3Wrapper_GUI ****
#AutoIt3Wrapper_Icon=..\AutoItv11.ico
#AutoIt3Wrapper_Res_Comment=Convert and resize from/to *NEW* WEBP, JPG, BMP, GIF, PNG,...
#AutoIt3Wrapper_Res_Description=Convert and resize from/to *NEW* WEBP, JPG, BMP, GIF, PNG,...
#AutoIt3Wrapper_Res_Fileversion=3.0.1.0
#AutoIt3Wrapper_Res_ProductName=Pics Converter V3
#AutoIt3Wrapper_Res_ProductVersion=3.0.1.0
#AutoIt3Wrapper_Res_CompanyName=cramaboule.com
#AutoIt3Wrapper_Run_Before=WriteTimestamp.exe "%in%"
#AutoIt3Wrapper_Run_After=copy %in% D:\Nextcloud\Cramy\Github\PicsConverter\
#AutoIt3Wrapper_Run_After=copy %out% D:\Nextcloud\Cramy\Github\PicsConverter\
#AutoIt3Wrapper_Run_After=copy ExtMsgBox.au3 D:\Nextcloud\Cramy\Github\PicsConverter\
#AutoIt3Wrapper_Run_After=copy StringSize.au3 D:\Nextcloud\Cramy\Github\PicsConverter\
#AutoIt3Wrapper_Run_Tidy=y
#Tidy_Parameters=/reel
#AutoIt3Wrapper_Run_Au3Stripper=n
#Au3Stripper_Parameters=/mo
#EndRegion ;**** Directives created by AutoIt3Wrapper_GUI ****
#Region ;Timestamp =====================
#                     2024/06/02 15:28:37
#EndRegion ;Timestamp =====================
#cs ----------------------------------------------------------------------------

	AutoIt Version: 3.3.16.1
	Author:         Cramaboule
	Date:			October 2009 V1

	Script Function: 	'Pics Converter V3 in (almoast) Pure AutoIt' made by Cramaboule Mai 2023
						Thanks to AdmiralAlkex for his help ! (on V1)

						Convert from/to JPG, BMP, GIF, PNG ,...!!! AND NEW WEBP
						Enjoy !

	Link: WebP: https://developers.google.com/speed/webp/download

	Bug:

	To Do:	WEBP: Resize webp, keep metadata.

	V3.0.1.0	02.06.2023:
				Improved: faster search with _ArrayConcatenate()
				Changed: -lossless can be used with -q (for WebP)
				Improved: faster convertion using -mt for WebP
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

$head = "Pics Conversion V3.0.1.0"

Local $Param = 0, $Decoder, $ToCombo, $ToComboOut, $OldOutEncoder, $Oldpxpercent, $Label2
Local $OldValSlider = '0', $OldJPGQuality = '100', $OldHeight, $Oldwidth, $OldCheckRatio, $OldLossless, $Parameter
Dim $aInterpolation[2][7] = [[$GDIP_INTERPOLATIONMODE_HIGHQUALITYBICUBIC, $GDIP_INTERPOLATIONMODE_HIGHQUALITYBILINEAR, $GDIP_INTERPOLATIONMODE_NEARESTNEIGHBOR, $GDIP_INTERPOLATIONMODE_BICUBIC, $GDIP_INTERPOLATIONMODE_BILINEAR, $GDIP_INTERPOLATIONMODE_HIGHQUALITY, $GDIP_INTERPOLATIONMODE_LOWQUALITY], ['Bicubic HQ (default)', 'Nearest neighbor', 'Bilinear HQ', 'Bicubic (low)', 'Bilinear (low)', 'High-quality', 'Low-quality']]

$pathWebP = _CheckWebP()

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

$Conv = GUICreate($head, 555, 210, -1, -1)
$Group1 = GUICtrlCreateGroup(" Input ", 5, 5, 140, 145)
$InputEncoder = GUICtrlCreateCombo("", 15, 120, 120, 25)
GUICtrlSetData(-1, $ToCombo)
$InputFolder = GUICtrlCreateInput("Input Folder", 15, 25, 120, 21)
$BrowseInput = GUICtrlCreateButton("Browse...", 60, 50, 75, 25, $WS_GROUP)
$Subfolder = GUICtrlCreateCheckbox("Subfolers ?", 60, 75, 75, 21)
$Label9 = GUICtrlCreateLabel("Convert from:", 15, 100, 67, 17)
GUICtrlCreateGroup("", -99, -99, 1, 1)
$Group4 = GUICtrlCreateGroup(" Resize ", 155, 5, 140, 145)
$Resizing = GUICtrlCreateCheckbox("Resizing", 165, 25, 55, 21)
$pxpercent = GUICtrlCreateCombo('px', 230, 25, 40, 25)
GUICtrlSetData(-1, '%')
$Ratio = GUICtrlCreateCheckbox("Keep aspect ratio", 170, 50, 115, 21)
GUICtrlSetState(-1, $GUI_CHECKED)
$Width = GUICtrlCreateInput("", 160, 75, 49, 21)
$Label4 = GUICtrlCreateLabel("px", 210, 80, 15, 17)
$Height = GUICtrlCreateInput("", 225, 75, 49, 21)
$Label5 = GUICtrlCreateLabel("px", 275, 80, 15, 17)
$Label3 = GUICtrlCreateLabel("Interpolation mode:", 165, 100, 104, 17)
$Interpolation = GUICtrlCreateCombo($aInterpolation[1][0], 165, 120, 120, 25, BitOR($CBS_DROPDOWN, $CBS_AUTOHSCROLL))
GUICtrlSetData(-1, $aInterpolation[1][1] & "|" & $aInterpolation[1][2] & "|" & $aInterpolation[1][3] & "|" & $aInterpolation[1][4] & "|" & $aInterpolation[1][5] & "|" & $aInterpolation[1][6])
GUICtrlSetState($Ratio, $GUI_DISABLE)
GUICtrlSetState($Width, $GUI_DISABLE)
GUICtrlSetState($Label4, $GUI_DISABLE)
GUICtrlSetState($Height, $GUI_DISABLE)
GUICtrlSetState($Label5, $GUI_DISABLE)
GUICtrlSetState($Label3, $GUI_DISABLE)
GUICtrlSetState($Interpolation, $GUI_DISABLE)
GUICtrlSetState($pxpercent, $GUI_DISABLE)
GUICtrlCreateGroup("", -99, -99, 1, 1)
$Group2 = GUICtrlCreateGroup(" Output ", 305, 5, 140, 145)
$OutputEncoder = GUICtrlCreateCombo("", 315, 120, 120, 25)
GUICtrlSetData(-1, $ToComboOut)
$OutputFolder = GUICtrlCreateInput("Output Folder", 315, 25, 120, 21)
$BrowseOutput = GUICtrlCreateButton("Browse...", 360, 50, 75, 25, $WS_GROUP)
$Label1 = GUICtrlCreateLabel("Convert to:", 315, 100, 56, 17)
GUICtrlCreateGroup("", -99, -99, 1, 1)
$Group3 = GUICtrlCreateGroup(" Quality ", 455, 5, 90, 145)
$Lossless = GUICtrlCreateCheckbox('Lossless', 465, 25, 60, 21)
$Slider = GUICtrlCreateSlider(495, 47, 35, 100, BitOR($TBS_VERT, $TBS_TOP, $TBS_LEFT))
$JPGQlty = GUICtrlCreateInput("100", 465, 87, 30, 21)
GUICtrlSetState($Group3, $GUI_ENABLE)
GUICtrlSetState($Slider, $GUI_DISABLE)
GUICtrlSetState($JPGQlty, $GUI_DISABLE)
GUICtrlSetState($Lossless, $GUI_DISABLE)
GUICtrlCreateGroup("", -99, -99, 1, 1)
$GO = GUICtrlCreateButton("Convert", 177, 160, 200, 40, $WS_GROUP)
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
				GUICtrlSetData($OutputFolder, $InFold)
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
					_CheckResize(0)
					GUICtrlSetState($Resizing, $GUI_UNCHECKED)
					GUICtrlSetState($Resizing, $GUI_DISABLE)
				EndIf
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
		Case $sPxpercent <> $Oldpxpercent
			If $sPxpercent = 'px' Then
				GUICtrlSetData($Label4, 'px')
				GUICtrlSetData($Label5, 'px')
				GUICtrlSetData($Width, '')
				GUICtrlSetData($Height, '')
			Else
				GUICtrlSetData($Label4, '%')
				GUICtrlSetData($Label5, '%')
				GUICtrlSetData($Width, '100')
				GUICtrlSetData($Height, '100')
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
		Case $nMsg = $GO
			$InPath = GUICtrlRead($InputFolder)
			$OutPath = GUICtrlRead($OutputFolder)
			$InEncoder = GUICtrlRead($InputEncoder)
			$OutEncoder = GUICtrlRead($OutputEncoder)
			If StringInStr($InPath, "\") = 0 Then
				_ExtMsgBox(16, 0, "Caution!", "Please select a folder!", 0, $Conv)
			ElseIf ($InEncoder = "") Or ($InEncoder = $OutEncoder And $OutEncoder <> "JPG" And Not (_IsChecked($Resizing))) Then
				_ExtMsgBox(16, 0, "Caution!", "Please choose different encoder/decoder or enable resizing", 0, $Conv)
			ElseIf GUICtrlRead($Width) = '' And GUICtrlRead($Height) = '' And _IsChecked($Resizing) Then
				_ExtMsgBox(16, 0, "Caution!", "Width and/or height cannot be empty", 0, $Conv)
			Else
;~ 				do the conversion process...
;~ 				do the progress bar GUI
				$aPos = WinGetPos($Conv)
				$Form1 = GUICreate("", 420, 100, ($aPos[0] + ($aPos[2] / 2)) - 210, ($aPos[1] + ($aPos[3] / 2)) - 70, BitOR($WS_POPUP, $WS_BORDER), $WS_EX_TOOLWINDOW, $Conv)
				Global $ProgFile = GUICtrlCreateProgress(10, 10, 400, 15, $PBS_SMOOTH)
				$Label2 = GUICtrlCreateLabel("", 10, 30, 400, 29)
				$ProgAll = GUICtrlCreateProgress(10, 60, 400, 15, $PBS_SMOOTH)
				$Label1 = GUICtrlCreateLabel("", 10, 80, 400, 17)
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
;~ 				==================================================== process itself
				Dim $FileList[1]
				$iFiles = _FindPathName($FileList, $InPath, "*." & $InEncoder, _IsChecked($Subfolder))
				If $iFiles <= 0 Then
					_ExtMsgBox(16, 0, "Caution!", "No files found or invalid path!", 0, $Conv)
				Else
					_GDIPlus_Startup()
					$clsid = _GDIPlus_EncodersGetCLSID($OutEncoder)
					For $i = 1 To $FileList[0]
						GUICtrlSetData($ProgFile, 0)
						GUICtrlSetData($ProgAll, ($i / $FileList[0]) * 100)
						$sPicsIn = $FileList[$i]
						$sTempPath = StringReplace($sPicsIn, "." & $InEncoder, "." & $OutEncoder) ; change extention
						$sPicsOut = StringReplace($sTempPath, $InPath, $OutPath)
						GUICtrlSetData($Label2, $sPicsIn & @CRLF & $sPicsOut)
						GUICtrlSetData($Label1, $i & " / " & $FileList[0])
						$aPath = StringSplit($sPicsOut, '\')
						$sPathOut = StringReplace($sPicsOut, $aPath[UBound($aPath) - 1], '') ; out path (without file name)
						If Not FileExists($sPathOut) Then DirCreate($sPathOut)
						If $InEncoder = 'WEBP' Or $OutEncoder = 'WEBP' Then
							If $OutEncoder = 'WEBP' Then
								_EncodeToWebP($pathWebP, $sPicsIn, $sPicsOut, _IsChecked($Lossless), GUICtrlRead($JPGQlty))
							ElseIf $InEncoder = 'WEBP' Then
								_DecodeFromWebP($pathWebP, $sPicsIn, $sPicsOut, $clsid)
							EndIf
						Else
							$hImage = _GDIPlus_ImageLoadFromFile($sPicsIn)
							GUICtrlSetData($ProgFile, 10)
							If _IsChecked($Resizing) Then ; resizing
								GUICtrlSetData($ProgFile, 20)
								$hImage = _Resize($hImage, _IsChecked($Ratio), GUICtrlRead($Height), GUICtrlRead($Width), GUICtrlRead($pxpercent), $iInterpolation)
								GUICtrlSetData($ProgFile, 40)
							EndIf
							GUICtrlSetData($ProgFile, 60)
							_GDIPlus_ImageSaveToFileEx($hImage, $sPicsOut, $clsid, $Param)
							_GDIPlus_ImageDispose($hImage)
						EndIf
						GUICtrlSetData($ProgFile, 80)
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
	If Not (FileExists('C:\Webp\cwebp.exe')) And Not (FileExists('C:\Webp\dwebp.exe')) Then
		DirCreate('C:\Webp')
		$Count += FileInstall('libwebp-1.4.0-windows-x64\bin\cwebp.exe', 'C:\Webp\cwebp.exe', $FC_OVERWRITE)
		$Count += FileInstall('libwebp-1.4.0-windows-x64\bin\dwebp.exe', 'C:\Webp\dwebp.exe', $FC_OVERWRITE)
	Else
		$Count = 2
	EndIf
	If $Count = 2 Then
		Return 'C:\Webp'
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

Func _EncodeToWebP($sspathWebP, $ssPicsIn, $ssPicsOut, $bbLossless, $sQuality)
	GUICtrlSetData($ProgFile, 10)
	$ssspathWebP = $sspathWebP & '\cwebp.exe'
	$ssPicsIn = '"' & $ssPicsIn & '"'
	$ssPicsOut = '-o "' & $ssPicsOut & '"'
	$sParameter = '-mt -quiet -q ' & $sQuality
	If _IsChecked($Lossless) Then
		$sParameter &= ' -lossless '
	EndIf
	$cmd = $ssspathWebP & ' ' & $sParameter & ' ' & $ssPicsIn & ' ' & $ssPicsOut
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
	_GDIPlus_ImageSaveToFileEx($hGdi, $ssPicsOut, $cclsid)
	_GDIPlus_BitmapDispose($hGdi)
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

Func _Resize($hhImage, $bRatio, $iiHeight, $iiWidth, $iipxpercent, $iInterpolation)
	$aDim = _GDIPlus_ImageGetDimension($hhImage)
	$iWidthIm = $aDim[0]
	$iHeightIm = $aDim[1]
	If $bRatio Then  ;ratio checked
		$fRatio = $iWidthIm / $iHeightIm  ; width / height
		If $iipxpercent = 'px' Then  ; px with ratio
			If $iiWidth <> '' Then
				$iWidth = $iiWidth
				$iHeight = Round($iiWidth / $fRatio)
			Else
				$iWidth = Round($iiHeight * $fRatio)
				$iHeight = $iiHeight
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
	$hhImage = _GDIPlus_ImageResize($hhImage, $iWidth, $iHeight, $iInterpolation) ;resize image
	Return $hhImage
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
	Else
		GUICtrlSetState($Ratio, $GUI_DISABLE)
		GUICtrlSetState($Width, $GUI_DISABLE)
		GUICtrlSetState($Label4, $GUI_DISABLE)
		GUICtrlSetState($Height, $GUI_DISABLE)
		GUICtrlSetState($Label5, $GUI_DISABLE)
		GUICtrlSetState($Label3, $GUI_DISABLE)
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
