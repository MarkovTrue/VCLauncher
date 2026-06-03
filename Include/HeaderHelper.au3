#include-once
#include <GDIPlus.au3>
#include <WinAPI.au3>

; ============================================================
; HeaderHelper.au3 — переиспользуемый рендер шапки окна
; ============================================================


; Рендерит шапку и устанавливает результат в GUI Pic-контрол.
; $hBitmapPrev передаётся ByRef для корректного освобождения старого HBITMAP.
Func _HeaderRenderToPic($iPicCtrl, ByRef $hBitmapPrev, $sIconsRootPath, $sTheme, $sFontName, $iDstW, $iDstH, ByRef $aHintRows, $iIconSizeDelta = 3, $nFontSize = 8.5, $iFontStyle = 1, $iBackColorArgb = -1, $sLeftIconPath = "", $sTitleApp = "", $sTitleVc = "", $sTitleRight = "", $sTitleVcVer = "")
	Local $hBmp = _HeaderCreateBitmap($sIconsRootPath, $sTheme, $sFontName, $iDstW, $iDstH, $aHintRows, $iIconSizeDelta, $nFontSize, $iFontStyle, $iBackColorArgb, $sLeftIconPath, $sTitleApp, $sTitleVc, $sTitleRight, $sTitleVcVer)
	If Not $hBmp Then Return False

	Local $hWnd = GUICtrlGetHandle($iPicCtrl)
	If Not $hWnd Then
		_WinAPI_DeleteObject($hBmp)
		Return False
	EndIf

	Local Const $STM_SETIMAGE = 0x0172, $IMAGE_BITMAP = 0
	Local $aRet = DllCall("user32.dll", "handle", "SendMessageW", "hwnd", $hWnd, _
			"uint", $STM_SETIMAGE, "wparam", $IMAGE_BITMAP, "lparam", $hBmp)
	If Not @error And IsArray($aRet) And $aRet[0] Then _WinAPI_DeleteObject($aRet[0])

	_HeaderDisposeBitmap($hBitmapPrev)
	$hBitmapPrev = $hBmp
	Return True
EndFunc   ;==>_HeaderRenderToPic


; Освобождает HBITMAP, если он был создан ранее.
Func _HeaderDisposeBitmap(ByRef $hBitmap)
	If $hBitmap Then
		_WinAPI_DeleteObject($hBitmap)
		$hBitmap = 0
	EndIf
EndFunc   ;==>_HeaderDisposeBitmap


; Создаёт HBITMAP шапки (иконка + два заголовка + подсказки), но не назначает его контролу.
; Раскладка: слева иконка и две строки заголовка, вертикальный разделитель,
; справа заголовок и две колонки подсказок. Фон — сплошной цвет темы, без тени текста.
Func _HeaderCreateBitmap($sIconsRootPath, $sTheme, $sFontName, $iDstW, $iDstH, ByRef $aHintRows, $iIconSizeDelta = 3, $nFontSize = 8.5, $iFontStyle = 1, $iBackColorArgb = -1, $sLeftIconPath = "", $sTitleApp = "", $sTitleVc = "", $sTitleRight = "", $sTitleVcVer = "")
	If Not IsArray($aHintRows) Then Return 0
	Local $iHintsCount = UBound($aHintRows)
	If $iHintsCount < 1 Then Return 0

	_GDIPlus_Startup()

	Local $hCanvas = _GDIPlus_BitmapCreateFromScan0($iDstW, $iDstH)
	Local $hGfx = _GDIPlus_ImageGetGraphicsContext($hCanvas)
	_GDIPlus_GraphicsSetInterpolationMode($hGfx, 7) ; HighQualityBicubic
	_GDIPlus_GraphicsSetTextRenderingHint($hGfx, 5) ; AntiAliasGridFit
	_GDIPlus_GraphicsSetSmoothingMode($hGfx, 2) ; HighQuality
	DllCall("gdiplus.dll", "int", "GdipSetPixelOffsetMode", "handle", $hGfx, "int", 2) ; HighQuality

	Local $bDarkSkin = (StringLower($sTheme) = "dark")
	If $iBackColorArgb = -1 Then $iBackColorArgb = ($bDarkSkin ? 0xFF1E1E1E : 0xFFF2F2F2)
	Local $hBackBrush = _GDIPlus_BrushCreateSolid($iBackColorArgb)
	_GDIPlus_GraphicsFillRect($hGfx, 0, 0, $iDstW, $iDstH, $hBackBrush)
	_GDIPlus_BrushDispose($hBackBrush)

	Local $iColText = $bDarkSkin ? 0xFFF0F0F0 : 0xFF1A1A1A
	Local $iColSub  = $bDarkSkin ? 0xFF969AA2 : 0xFF6E6E6E
	Local $iColSep  = $bDarkSkin ? 0xFF484C54 : 0xFFD0D0D0
	Local $sIconTheme = $bDarkSkin ? "Light" : "Dark"

	; --- логотип слева, прижат к верху ---
	Local Const $iIcoTarget = 36
	Local $iIcoX = 14
	Local $iIcoY = 14
	If $sLeftIconPath <> "" And FileExists($sLeftIconPath) Then
		Local $hIconBmp = __HeaderLoadHeaderIcon($sLeftIconPath, $iIcoTarget)
		If $hIconBmp Then
			__HeaderDrawImageAlpha($hGfx, $hIconBmp, $iIcoX, $iIcoY, $iIcoTarget, $iIcoTarget, 1.0)
			_GDIPlus_ImageDispose($hIconBmp)
		EndIf
	EndIf

	; --- шрифты и кисти ---
	Local $hFamily = _GDIPlus_FontFamilyCreate($sFontName)
	Local $hFont = _GDIPlus_FontCreate($hFamily, $nFontSize, $iFontStyle)
	Local $hTitleFont = _GDIPlus_FontCreate($hFamily, 10, 1) ; заголовок приложения, bold
	Local $hVcFont = _GDIPlus_FontCreate($hFamily, 10, 0)    ; "Video-compare" + версия
	Local $hRightTitleFont = _GDIPlus_FontCreate($hFamily, 10, 1) ; заголовок колонки клавиш
	Local $hBrush = _GDIPlus_BrushCreateSolid($iColText)
	Local $hSubBrush = _GDIPlus_BrushCreateSolid($iColSub)
	Local $hFormat = _GDIPlus_StringFormatCreate()

	; --- заголовок под иконкой: VCLauncher (жирный), затем Video-compare и версия
	; двумя строками (обычный, приглушённый). Блок прижат к левому краю под иконкой. ---
	Local $iTitleX = $iIcoX
	Local $iTitleY = $iIcoY + $iIcoTarget + 6
	Local $iLineStep = 17
	__HeaderDrawText($hGfx, $sTitleApp, $hTitleFont, $hBrush, $hFormat, $iTitleX, $iTitleY, 130)
	__HeaderDrawText($hGfx, $sTitleVc, $hVcFont, $hSubBrush, $hFormat, $iTitleX, $iTitleY + $iLineStep, 130)
	__HeaderDrawText($hGfx, $sTitleVcVer, $hVcFont, $hSubBrush, $hFormat, $iTitleX, $iTitleY + $iLineStep * 2, 130)

	; --- вертикальный разделитель (заголовок теперь под иконкой, блок ужат влево) ---
	Local $iSepX = 148
	Local $hPen = _GDIPlus_PenCreate($iColSep, 1)
	_GDIPlus_GraphicsDrawLine($hGfx, $iSepX, 14, $iSepX, $iDstH - 14, $hPen)
	_GDIPlus_PenDispose($hPen)

	; --- правый заголовок ---
	Local $iRightX = 162
	__HeaderDrawText($hGfx, $sTitleRight, $hRightTitleFont, $hBrush, $hFormat, $iRightX, 11, $iDstW - $iRightX - 8)

	; --- подсказки: две ровные колонки под правым заголовком ---
	; Ширина колонок адаптивна: половина зоны от $iRightX до правого края.
	; Так вторая колонка ($iRightX + $iColGap) не выходит за край при любой ширине окна.
	Local $iRowsPerCol = Int(($iHintsCount + 1) / 2)
	Local $iColGap = Int(($iDstW - $iRightX - 8) / 2)
	Local $iColW = $iColGap - 8
	Local $iKeysY0 = 36
	Local $iLineH = 15
	Local $iBottomPad = 6
	Local $nRowStep = $iLineH + 3
	If $iRowsPerCol > 1 Then $nRowStep = ($iDstH - $iKeysY0 - $iLineH - $iBottomPad) / ($iRowsPerCol - 1)

	For $i = 0 To $iRowsPerCol - 1
		Local $iRowY = $iKeysY0 + Int($i * $nRowStep)
		If $i < $iHintsCount Then
			__HeaderDrawHotkeyHintRowLocalized($hGfx, $sIconsRootPath, $sIconTheme, $aHintRows[$i], $hFont, $hFormat, $hBrush, $iRightX, $iRowY, $iColW, $iLineH, $iIconSizeDelta)
		EndIf
		Local $iRight = $i + $iRowsPerCol
		If $iRight < $iHintsCount Then
			__HeaderDrawHotkeyHintRowLocalized($hGfx, $sIconsRootPath, $sIconTheme, $aHintRows[$iRight], $hFont, $hFormat, $hBrush, $iRightX + $iColGap, $iRowY, $iColW, $iLineH, $iIconSizeDelta)
		EndIf
	Next

	_GDIPlus_FontDispose($hFont)
	_GDIPlus_FontDispose($hTitleFont)
	_GDIPlus_FontDispose($hVcFont)
	_GDIPlus_FontDispose($hRightTitleFont)
	_GDIPlus_FontFamilyDispose($hFamily)
	_GDIPlus_BrushDispose($hBrush)
	_GDIPlus_BrushDispose($hSubBrush)
	_GDIPlus_StringFormatDispose($hFormat)

	Local $hBmp = _GDIPlus_BitmapCreateHBITMAPFromBitmap($hCanvas)
	_GDIPlus_GraphicsDispose($hGfx)
	_GDIPlus_BitmapDispose($hCanvas)
	_GDIPlus_Shutdown()
	Return $hBmp
EndFunc   ;==>_HeaderCreateBitmap


; Рисует строку текста в заданной точке (левое выравнивание, фикс. ширина layout).
Func __HeaderDrawText($hGfx, $sText, $hFont, $hBrush, $hFormat, $iX, $iY, $iW = 300)
	If $sText = "" Then Return
	Local $tLayout = _GDIPlus_RectFCreate($iX, $iY, $iW, 22)
	_GDIPlus_GraphicsDrawStringEx($hGfx, $sText, $hFont, $tLayout, $hFormat, $hBrush)
EndFunc   ;==>__HeaderDrawText


Func __HeaderDrawHotkeyHintRowLocalized($hGfx, $sIconsRootPath, $sTheme, $sRawLine, $hTextFont, $hFormat, $hBrush, $iX, $iY, $iW, $iH, $iIconSizeDelta)
	Local $sIcons = __HeaderHotkeyIconsFromTags($sRawLine)
	Local $sText = __HeaderHotkeyTextFromTags($sRawLine)
	__HeaderDrawHotkeyHintRow($hGfx, $sIconsRootPath, $sTheme, $sIcons, $sText, $hTextFont, $hFormat, $hBrush, $iX, $iY, $iW, $iH, $iIconSizeDelta)
EndFunc   ;==>__HeaderDrawHotkeyHintRowLocalized


Func __HeaderHotkeyIconsFromTags($sLine)
	Local $aTags = StringRegExp($sLine, "\{([A-Za-z0-9_]+)\}", 3)
	If @error Or Not IsArray($aTags) Then Return ""

	Local $sIcons = ""
	For $sTag In $aTags
		If $sIcons <> "" Then $sIcons &= "+"
		$sIcons &= StringUpper($sTag)
	Next
	Return $sIcons
EndFunc   ;==>__HeaderHotkeyIconsFromTags


Func __HeaderHotkeyTextFromTags($sLine)
	Local $sText = StringRegExpReplace($sLine, "\{[A-Za-z0-9_]+\}", "")
	Return StringStripWS($sText, 3)
EndFunc   ;==>__HeaderHotkeyTextFromTags


Func __HeaderDrawHotkeyHintRow($hGfx, $sIconsRootPath, $sTheme, $sIcons, $sText, $hTextFont, $hFormat, $hBrush, $iX, $iY, $iW, $iH, $iIconSizeDelta)
	Local $aTokens = StringSplit($sIcons, "+", 2)
	Local $iCurX = $iX
	Local $iIconSize = 14
	If $iIconSizeDelta > 0 Then $iIconSize += Int($iIconSizeDelta / 2)

	For $vToken In $aTokens
		Local $sIconPath = __HeaderResolveIconPath($sIconsRootPath, $sTheme, $vToken)
		If $sIconPath = "" Or Not FileExists($sIconPath) Then ContinueLoop

		Local $hIcon = _GDIPlus_ImageLoadFromFile($sIconPath)
		If @error Or Not $hIcon Then ContinueLoop

		Local $iIconY = $iY + Int(($iH - $iIconSize) / 2)
		_GDIPlus_GraphicsDrawImageRect($hGfx, $hIcon, $iCurX, $iIconY, $iIconSize, $iIconSize)
		_GDIPlus_ImageDispose($hIcon)

		$iCurX += $iIconSize + 1
	Next

	Local $iTextX = $iCurX + 1
	Local $iTextW = $iW - ($iTextX - $iX)
	If $iTextW < 8 Then Return

	; Текст без тени (полупрозрачная подложка убрана).
	Local $tLayout = _GDIPlus_RectFCreate($iTextX, $iY - 1, $iTextW, $iH)
	_GDIPlus_GraphicsDrawStringEx($hGfx, $sText, $hTextFont, $tLayout, $hFormat, $hBrush)
EndFunc   ;==>__HeaderDrawHotkeyHintRow


Func __HeaderResolveIconPath($sIconsRootPath, $sTheme, $sToken)
	Local $sStyleFolder = "Light"
	Local $sFallbackFolder = "Default"
	If StringLower($sTheme) = "dark" Then
		$sStyleFolder = "Dark"
		$sFallbackFolder = "Double"
	EndIf

	Local $sFile = ""
	Switch StringUpper($sToken)
		Case "H"
			$sFile = "KeyboardH"
		Case "Y"
			$sFile = "KeyboardY"
		Case "0"
			$sFile = "Keyboard0"
		Case "3"
			$sFile = "Keyboard3"
		Case "4"
			$sFile = "Keyboard4"
		Case "5"
			$sFile = "Keyboard5"
		Case "6"
			$sFile = "Keyboard6"
		Case "7"
			$sFile = "Keyboard7"
		Case "8"
			$sFile = "Keyboard8"
		Case "9"
			$sFile = "Keyboard9"
		Case "LEFT"
			$sFile = "KeyboardArrowLeft"
		Case "RIGHT"
			$sFile = "KeyboardArrowRight"
		Case "UP"
			$sFile = "KeyboardArrowUp"
		Case "DOWN"
			$sFile = "KeyboardArrowDown"
		Case "PGUP"
			$sFile = "KeyboardPageUp"
		Case "PGDOWN"
			$sFile = "KeyboardPageDown"
		Case "PLUS"
			$sFile = "KeyboardPlus"
		Case "MINUS"
			$sFile = "KeyboardMinus"
		Case "CTRL"
			$sFile = "KeyboardCtrl"
		Case "ALT"
			$sFile = "KeyboardAlt"
		Case Else
			Return ""
	EndSwitch

	Local $sPrimaryPath = $sIconsRootPath & "\\" & $sStyleFolder & "\\" & $sFile & ".png"
	If FileExists($sPrimaryPath) Then Return $sPrimaryPath

	; Обратная совместимость со старыми папками набора
	Local $sFallbackPath = $sIconsRootPath & "\\" & $sFallbackFolder & "\\" & $sFile & ".png"
	If FileExists($sFallbackPath) Then Return $sFallbackPath

	Return $sPrimaryPath
EndFunc   ;==>__HeaderResolveIconPath


; Рисует GDI+-изображение с заданной альфой через ColorMatrix.
Func __HeaderDrawImageAlpha($hGfx, $hImage, $iX, $iY, $iW, $iH, $nAlpha)
	Local $tCM = DllStructCreate("float[25]")
	DllStructSetData($tCM, 1, 1.0, 1)  ; [0][0]
	DllStructSetData($tCM, 1, 1.0, 7)  ; [1][1]
	DllStructSetData($tCM, 1, 1.0, 13) ; [2][2]
	DllStructSetData($tCM, 1, $nAlpha, 19) ; [3][3]
	DllStructSetData($tCM, 1, 1.0, 25) ; [4][4]

	Local $aAttr = DllCall("gdiplus.dll", "int", "GdipCreateImageAttributes", "ptr*", 0)
	If @error Or Not IsArray($aAttr) Or $aAttr[0] <> 0 Then Return False
	Local $hAttr = $aAttr[1]

	Local Const $ColorAdjustTypeBitmap = 1
	DllCall("gdiplus.dll", "int", "GdipSetImageAttributesColorMatrix", _
			"ptr", $hAttr, "int", $ColorAdjustTypeBitmap, "bool", True, _
			"ptr", DllStructGetPtr($tCM), "ptr", 0, "int", 0)

	; Clamp по краям — без полупрозрачной каймы при бикубическом уменьшении.
	Local Const $WrapModeTileFlipXY = 3
	DllCall("gdiplus.dll", "int", "GdipSetImageAttributesWrapMode", _
			"ptr", $hAttr, "int", $WrapModeTileFlipXY, "uint", 0, "bool", False)

	; Исходный прямоугольник — натуральный размер изображения; целевой — заданный.
	; Масштаб выполняется с InterpolationMode контекста (HighQualityBicubic).
	Local $iSrcW = _GDIPlus_ImageGetWidth($hImage)
	Local $iSrcH = _GDIPlus_ImageGetHeight($hImage)

	Local Const $UnitPixel = 2
	DllCall("gdiplus.dll", "int", "GdipDrawImageRectRectI", _
			"handle", $hGfx, "handle", $hImage, _
			"int", $iX, "int", $iY, "int", $iW, "int", $iH, _
			"int", 0, "int", 0, "int", $iSrcW, "int", $iSrcH, _
			"int", $UnitPixel, "ptr", $hAttr, "ptr", 0, "ptr", 0)

	DllCall("gdiplus.dll", "int", "GdipDisposeImageAttributes", "ptr", $hAttr)
	Return True
EndFunc   ;==>__HeaderDrawImageAlpha


; Загружает логотип шапки. PNG читается напрямую через GDI+ (полная 8-бит альфа,
; сглаженные края); .ico — через кадр нужного размера. Масштаб делает вызывающий код.
Func __HeaderLoadHeaderIcon($sPath, $iTargetSize = 48)
	If StringRight(StringLower($sPath), 4) = ".png" Then
		Local $hImg = _GDIPlus_ImageLoadFromFile($sPath)
		If @error Or Not $hImg Then Return 0
		Return $hImg
	EndIf
	Return __HeaderLoadIcoBitmap($sPath, $iTargetSize)
EndFunc   ;==>__HeaderLoadHeaderIcon


; Загружает из .ico кадр заданного целевого размера и возвращает GDI+ Bitmap.
; Рисуется 1:1, поэтому размер фиксирован и не зависит от крупных кадров в .ico
; (256px и т.п. нужны для иконки exe/проводника, но в шапке были бы огромными).
Func __HeaderLoadIcoBitmap($sIcoPath, $iTargetSize = 48)
	Local $iMax = __HeaderIcoMaxSize($sIcoPath)
	Local $iLoad = $iTargetSize
	If $iMax > 0 And $iMax < $iTargetSize Then $iLoad = $iMax ; не растягивать сверх доступного кадра

	Local Const $IMAGE_ICON = 1
	Local Const $LR_LOADFROMFILE = 0x00000010
	Local $aLoad = DllCall("user32.dll", "handle", "LoadImageW", _
			"ptr", 0, "wstr", $sIcoPath, "uint", $IMAGE_ICON, _
			"int", $iLoad, "int", $iLoad, "uint", $LR_LOADFROMFILE)
	If @error Or Not IsArray($aLoad) Or $aLoad[0] = 0 Then Return 0

	Local $hIcon = $aLoad[0]
	Local $aCreate = DllCall("gdiplus.dll", "int", "GdipCreateBitmapFromHICON", "handle", $hIcon, "ptr*", 0)
	DllCall("user32.dll", "bool", "DestroyIcon", "handle", $hIcon)
	If @error Or Not IsArray($aCreate) Or $aCreate[0] <> 0 Then Return 0
	Return $aCreate[2]
EndFunc   ;==>__HeaderLoadIcoBitmap


; Разбирает заголовок .ico и возвращает максимальную ширину среди записей.
Func __HeaderIcoMaxSize($sIcoPath)
	Local $hFile = FileOpen($sIcoPath, 16) ; binary
	If $hFile = -1 Then Return 0

	Local $bHeader = FileRead($hFile, 6)
	If @error Or BinaryLen($bHeader) < 6 Then
		FileClose($hFile)
		Return 0
	EndIf

	Local $iCount = Dec(Hex(BinaryMid($bHeader, 5, 1)), 2) + Dec(Hex(BinaryMid($bHeader, 6, 1)), 2) * 256
	Local $iMax = 0
	For $i = 1 To $iCount
		Local $bEntry = FileRead($hFile, 16)
		If @error Or BinaryLen($bEntry) < 2 Then ExitLoop
		Local $iW = Dec(Hex(BinaryMid($bEntry, 1, 1)), 2)
		If $iW = 0 Then $iW = 256
		If $iW > $iMax Then $iMax = $iW
	Next
	FileClose($hFile)
	Return $iMax
EndFunc   ;==>__HeaderIcoMaxSize


Func __HeaderGetHotkeyGlyph($sToken)
	Switch StringUpper($sToken)
		Case "H"
			Return ChrW(0xE085)
		Case "Y"
			Return ChrW(0xE0DE)
		Case "0"
			Return ChrW(0xE002)
		Case "3"
			Return ChrW(0xE008)
		Case "4"
			Return ChrW(0xE00A)
		Case "5"
			Return ChrW(0xE00C)
		Case "6"
			Return ChrW(0xE00E)
		Case "7"
			Return ChrW(0xE010)
		Case "8"
			Return ChrW(0xE012)
		Case "9"
			Return ChrW(0xE014)
		Case "LEFT"
			Return ChrW(0xE020)
		Case "RIGHT"
			Return ChrW(0xE022)
		Case "UP"
			Return ChrW(0xE024)
		Case "DOWN"
			Return ChrW(0xE01E)
		Case "PGUP"
			Return ChrW(0xE0A8)
		Case "PGDOWN"
			Return ChrW(0xE0A6)
		Case "PLUS"
			Return ChrW(0xE0AC)
		Case "MINUS"
			Return ChrW(0xE095)
		Case "CTRL"
			Return ChrW(0xE048)
		Case "ALT"
			Return ChrW(0xE040)
		Case Else
			Return ""
	EndSwitch
EndFunc   ;==>__HeaderGetHotkeyGlyph


