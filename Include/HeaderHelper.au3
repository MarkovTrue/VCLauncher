#include-once
#include <GDIPlus.au3>
#include <WinAPI.au3>

Global $g_hHeaderPrivateFontCollection = 0
Global Const $gc_sHeaderKeyFontFamily = "Kenney Input Keyboard & Mouse"

; ============================================================
; HeaderHelper.au3 — переиспользуемый рендер шапки окна
; ============================================================


; Рендерит шапку и устанавливает результат в GUI Pic-контрол.
; $hBitmapPrev передаётся ByRef для корректного освобождения старого HBITMAP.
Func _HeaderRenderToPic($iPicCtrl, ByRef $hBitmapPrev, $sLogoPath, $sIconsRootPath, $sTheme, $sFontName, $iDstW, $iDstH, ByRef $aHintRows, $iIconSizeDelta = 3, $nFontSize = 8.5, $iFontStyle = 1, $iBackColorArgb = -1)
	Local $hBmp = _HeaderCreateBitmap($sLogoPath, $sIconsRootPath, $sTheme, $sFontName, $iDstW, $iDstH, $aHintRows, $iIconSizeDelta, $nFontSize, $iFontStyle, $iBackColorArgb)
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


; Создаёт HBITMAP шапки (логотип + подсказки), но не назначает его контролу.
Func _HeaderCreateBitmap($sLogoPath, $sIconsRootPath, $sTheme, $sFontName, $iDstW, $iDstH, ByRef $aHintRows, $iIconSizeDelta = 3, $nFontSize = 8.5, $iFontStyle = 1, $iBackColorArgb = -1)
	If Not FileExists($sLogoPath) Then Return 0
	If Not IsArray($aHintRows) Then Return 0

	Local $iHintsCount = UBound($aHintRows)
	If $iHintsCount < 1 Then Return 0

	_GDIPlus_Startup()
	Local $hImage = _GDIPlus_ImageLoadFromFile($sLogoPath)
	If @error Or Not $hImage Then
		_GDIPlus_Shutdown()
		Return 0
	EndIf

	Local $iSrcW = _GDIPlus_ImageGetWidth($hImage)
	Local $iSrcH = _GDIPlus_ImageGetHeight($hImage)

	Local $hCanvas = _GDIPlus_BitmapCreateFromScan0($iDstW, $iDstH)
	Local $hGfx = _GDIPlus_ImageGetGraphicsContext($hCanvas)
	_GDIPlus_GraphicsSetInterpolationMode($hGfx, 7) ; HighQualityBicubic
	_GDIPlus_GraphicsSetTextRenderingHint($hGfx, 5) ; AntiAliasGridFit

	Local $bDarkSkin = (StringLower($sTheme) = "dark")
	If $iBackColorArgb = -1 Then $iBackColorArgb = ($bDarkSkin ? 0xFF121212 : 0xFFF2F2F2)
	Local $hBackBrush = _GDIPlus_BrushCreateSolid($iBackColorArgb)
	_GDIPlus_GraphicsFillRect($hGfx, 0, 0, $iDstW, $iDstH, $hBackBrush)
	_GDIPlus_BrushDispose($hBackBrush)

	; Логотип рисуется 1:1 без растяжения.
	_GDIPlus_GraphicsDrawImageRect($hGfx, $hImage, 0, 0, $iSrcW, $iSrcH)
	_GDIPlus_ImageDispose($hImage)

	__HeaderEnsureKeyFontCollection()

	Local $hFamily = _GDIPlus_FontFamilyCreate($sFontName)
	Local $hFont = _GDIPlus_FontCreate($hFamily, $nFontSize, $iFontStyle)
	Local $hKeyFamily = __HeaderCreateKeyFontFamily()
	If @error Or Not $hKeyFamily Then $hKeyFamily = $hFamily
	Local $hKeyFont = _GDIPlus_FontCreate($hKeyFamily, 26, $iFontStyle)
	Local $hBrush = _GDIPlus_BrushCreateSolid($bDarkSkin ? 0xFFF5F5F5 : 0xFF111111)
	Local $hBrushShadow = _GDIPlus_BrushCreateSolid($bDarkSkin ? 0xAA000000 : 0xAAFFFFFF)
	Local $sIconTheme = $bDarkSkin ? "Light" : "Dark"
	Local $hFormat = _GDIPlus_StringFormatCreate()

	Local $iRowsPerCol = Int(($iHintsCount + 1) / 2)
	Local $iCol1X = Int($iDstW * 0.2) - 12
	Local $iColW = Int(($iDstW - $iCol1X - 8) / 2)
	Local $iCol2X = $iCol1X + $iColW

	; Шахматный порядок в пределах двух колонок: X-сдвиг по строкам
	; + небольшой вертикальный сдвиг правой колонки.
	Local $iTopPad = 4
	Local $iBottomPad = 4
	Local $iAvailH = $iDstH - $iTopPad - $iBottomPad
	Local $iLineH = 16
	If $iAvailH < $iLineH Then $iLineH = $iAvailH

	Local $iRowStep = 0
	If $iRowsPerCol > 1 Then
		$iRowStep = Int(($iAvailH - $iLineH) / ($iRowsPerCol - 1))
		If $iRowStep < 11 Then $iRowStep = 11
	EndIf

	Local $iBlockH = (($iRowsPerCol - 1) * $iRowStep) + $iLineH
	Local $iTextY = Int(($iDstH - $iBlockH) / 2)
	If $iTextY < $iTopPad Then $iTextY = $iTopPad

	Local $iChessShiftX = 8
	Local $iRightYOffset = 0 ; убираем вертикальный сдвиг для выравнивания колонок
	Local $iDrawW = $iColW - $iChessShiftX - 2
	If $iDrawW < 40 Then $iDrawW = $iColW - 2

	For $i = 0 To $iRowsPerCol - 1
		If $i < $iHintsCount Then
			Local $iLeftX = $iCol1X + (Mod($i, 2) = 1 ? $iChessShiftX : 0)
			Local $iLeftY = $iTextY + ($i * $iRowStep)
			__HeaderDrawHotkeyHintRowLocalized($hGfx, $sIconsRootPath, $sIconTheme, $aHintRows[$i], $hKeyFont, $hFont, $hFormat, $hBrushShadow, $hBrush, $iLeftX, $iLeftY, $iDrawW, $iLineH, $iIconSizeDelta)
		EndIf

		Local $iRight = $i + $iRowsPerCol
		If $iRight < $iHintsCount Then
			   Local $iRightX = $iCol2X + (Mod($i, 2) = 0 ? $iChessShiftX : 0)
			   Local $iRightY = $iTextY + ($i * $iRowStep) ; без сдвига
			If ($iRightY + $iLineH) > ($iDstH - $iBottomPad) Then $iRightY = $iDstH - $iBottomPad - $iLineH
			__HeaderDrawHotkeyHintRowLocalized($hGfx, $sIconsRootPath, $sIconTheme, $aHintRows[$iRight], $hKeyFont, $hFont, $hFormat, $hBrushShadow, $hBrush, $iRightX, $iRightY, $iDrawW, $iLineH, $iIconSizeDelta)
		EndIf
	Next

	_GDIPlus_FontDispose($hKeyFont)
	If $hKeyFamily <> $hFamily Then _GDIPlus_FontFamilyDispose($hKeyFamily)
	_GDIPlus_FontDispose($hFont)
	_GDIPlus_FontFamilyDispose($hFamily)
	_GDIPlus_BrushDispose($hBrushShadow)
	_GDIPlus_BrushDispose($hBrush)
	_GDIPlus_StringFormatDispose($hFormat)

	Local $hBmp = _GDIPlus_BitmapCreateHBITMAPFromBitmap($hCanvas)
	_GDIPlus_GraphicsDispose($hGfx)
	_GDIPlus_BitmapDispose($hCanvas)
	_GDIPlus_Shutdown()
	Return $hBmp
EndFunc   ;==>_HeaderCreateBitmap


Func __HeaderDrawHotkeyHintRowLocalized($hGfx, $sIconsRootPath, $sTheme, $sRawLine, $hKeyFont, $hTextFont, $hFormat, $hBrushShadow, $hBrush, $iX, $iY, $iW, $iH, $iIconSizeDelta)
	Local $sIcons = __HeaderHotkeyIconsFromTags($sRawLine)
	Local $sText = __HeaderHotkeyTextFromTags($sRawLine)
	__HeaderDrawHotkeyHintRow($hGfx, $sIconsRootPath, $sTheme, $sIcons, $sText, $hKeyFont, $hTextFont, $hFormat, $hBrushShadow, $hBrush, $iX, $iY, $iW, $iH, $iIconSizeDelta)
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


Func __HeaderDrawHotkeyHintRow($hGfx, $sIconsRootPath, $sTheme, $sIcons, $sText, $hKeyFont, $hTextFont, $hFormat, $hBrushShadow, $hBrush, $iX, $iY, $iW, $iH, $iIconSizeDelta)
	#forceref $hKeyFont
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

	Local $tLayoutShadow = _GDIPlus_RectFCreate($iTextX + 1, $iY, $iTextW, $iH)
	Local $tLayout = _GDIPlus_RectFCreate($iTextX, $iY - 1, $iTextW, $iH)
	_GDIPlus_GraphicsDrawStringEx($hGfx, $sText, $hTextFont, $tLayoutShadow, $hFormat, $hBrushShadow)
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


Func __HeaderEnsureKeyFontCollection()
	If $g_hHeaderPrivateFontCollection Then Return True

	Local $sFontPath = @ScriptDir & "\Assets\Fonts\kenney_input_keyboard_&_mouse.ttf"
	If Not FileExists($sFontPath) Then Return False

	; Загружаем шрифт в приватную коллекцию GDI+, доступную только процессу.
	Local $aNew = DllCall("gdiplus.dll", "int", "GdipNewPrivateFontCollection", "ptr*", 0)
	If @error Or Not IsArray($aNew) Or $aNew[0] <> 0 Or $aNew[1] = 0 Then Return False

	Local $hCollection = $aNew[1]
	Local $aAdd = DllCall("gdiplus.dll", "int", "GdipPrivateAddFontFile", "ptr", $hCollection, "wstr", $sFontPath)
	If @error Or Not IsArray($aAdd) Or $aAdd[0] <> 0 Then
		Return False
	EndIf

	$g_hHeaderPrivateFontCollection = $hCollection
	Return True
EndFunc   ;==>__HeaderEnsureKeyFontCollection


Func __HeaderCreateKeyFontFamily()
	If __HeaderEnsureKeyFontCollection() Then
		Local $aCreate = DllCall("gdiplus.dll", "int", "GdipCreateFontFamilyFromName", _
				"wstr", $gc_sHeaderKeyFontFamily, _
				"ptr", $g_hHeaderPrivateFontCollection, _
				"ptr*", 0)
		If Not @error And IsArray($aCreate) And $aCreate[0] = 0 And $aCreate[3] <> 0 Then Return $aCreate[3]
	EndIf

	; Fallback на системный поиск, если приватная загрузка недоступна.
	Return _GDIPlus_FontFamilyCreate($gc_sHeaderKeyFontFamily)
EndFunc   ;==>__HeaderCreateKeyFontFamily


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


