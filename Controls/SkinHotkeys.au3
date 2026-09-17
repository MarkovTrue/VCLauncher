#include-once
#include "SkinCore.au3"
#include <StaticConstants.au3>
#include <WindowsConstants.au3>
#include <Math.au3>
#include <StringConstants.au3>

; ============================================================
; SkinHotkeys.au3 – раскрывающаяся шпаргалка горячих клавиш
; ============================================================
; Карточка под шапкой: заголовок и $gc_iSkinHkCols колонок строк «клавиши – действие».
; Клавиша – готовая иконка из Assets\KeyIcons (Light\Dark), для токена без иконки
; запасной бейдж. Показывает и прячет панель вызывающий, высоту заранее даёт
; _SkinHotkeysHeight.
;
; Отдельные Pic, чтобы ресайз окна перерисовывал только рамку:
;   - фон с рамкой карточки тянется с окном;
;   - заголовок по ширине текста стоит на месте (в фоне дёргался при ресайзе);
;   - колонка строк – картинка под свой контент. На опорной ширине карточки
;     колонки разнесены по ней с равными зазорами, шире – едут с центром своей трети.
; Колонки не ужимаются и обязаны влезть при минимальной ширине окна.

Global Const $gc_iSkinHkRowH = 19    ; шаг строки
Global Const $gc_iSkinHkPad = 14     ; поле панели
Global Const $gc_iSkinHkTitleH = 30  ; заголовок панели
Global Const $gc_iSkinHkIconSize = 16 ; сторона готовой иконки клавиши
Global Const $gc_iSkinHkColGap = 20  ; зазор между колонками
Global Const $gc_iSkinHkKeyGap = 8   ; зазор от ряда клавиш до подписи
; Сдвиг подписи вниз: GDI+ центрирует строку вместе с нижними выносными,
; и заглавные вставали на 2 px выше центра иконки (замер по снимку).
Global Const $gc_iSkinHkTextDY = 2
Global Const $gc_iSkinHkCols = 3     ; число колонок строк

; Реестр панелей. Колонки:
; 0 ControlID фон, 1 HBITMAP фона, 2 W фона, 3 H,
; 4 заголовок, 5 строки [n][2],
; 6 ControlID заголовка, 7 HBITMAP заголовка,
; 8 опорная W фона (при создании, от неё считается сдвиг колонок),
; дальше по тройке на колонку строк (начало – __SkinHotkeysColBase):
; +0 ControlID, +1 HBITMAP, +2 W по контенту (0 – колонка пустая)
Global Const $gc_iSkinHkColBase = 9
Global $g_aSkinHotkeys[0][$gc_iSkinHkColBase + 3 * $gc_iSkinHkCols]

; Кэш готовых картинок клавиш из Assets\KeyIcons: ключ "ТОКЕН|Light|Dark" -> GDI+ Image
Global $g_mSkinHkIcons[]


; Высота панели под $iRows строк, разложенных по $gc_iSkinHkCols колонкам
Func _SkinHotkeysHeight($iRows)
	Local $iPerCol = Ceiling($iRows / $gc_iSkinHkCols)
	Return $gc_iSkinHkTitleH + $iPerCol * $gc_iSkinHkRowH + $gc_iSkinHkPad
EndFunc   ;==>_SkinHotkeysHeight


; $aRows – [n][2]: клавиши через "|" и подпись. Строки заполняют колонки поровну,
; сверху вниз и слева направо. ControlID фона – «хендл» панели для функций модуля.
Func _SkinHotkeysCreate($iX, $iY, $iW, $sTitle, ByRef $aRows)
	Local $iH = _SkinHotkeysHeight(UBound($aRows))

	Local $iCtrlBg = GUICtrlCreatePic("", $iX, $iY, $iW, $iH, $WS_CLIPSIBLINGS)
	_SkinDockFixed($iCtrlBg)

	Local $iIndex = UBound($g_aSkinHotkeys)
	ReDim $g_aSkinHotkeys[$iIndex + 1][UBound($g_aSkinHotkeys, 2)]
	$g_aSkinHotkeys[$iIndex][0] = $iCtrlBg
	$g_aSkinHotkeys[$iIndex][1] = 0
	$g_aSkinHotkeys[$iIndex][2] = $iW
	$g_aSkinHotkeys[$iIndex][3] = $iH
	$g_aSkinHotkeys[$iIndex][4] = $sTitle
	$g_aSkinHotkeys[$iIndex][5] = $aRows
	$g_aSkinHotkeys[$iIndex][6] = GUICtrlCreatePic("", $iX, $iY, 1, 1)
	$g_aSkinHotkeys[$iIndex][7] = 0
	$g_aSkinHotkeys[$iIndex][8] = $iW
	_SkinDockFixed($g_aSkinHotkeys[$iIndex][6])
	For $c = 0 To $gc_iSkinHkCols - 1
		Local $iBase = __SkinHotkeysColBase($c)
		$g_aSkinHotkeys[$iIndex][$iBase] = GUICtrlCreatePic("", $iX, $iY, 1, 1)
		$g_aSkinHotkeys[$iIndex][$iBase + 1] = 0
		$g_aSkinHotkeys[$iIndex][$iBase + 2] = 0
		_SkinDockFixed($g_aSkinHotkeys[$iIndex][$iBase])
	Next

	__SkinHotkeysRenderTitle($iIndex)
	__SkinHotkeysRenderColumns($iIndex)
	__SkinHotkeysRenderBg($iIndex) ; последним: расставляет колонки по их ширине
	Return $iCtrlBg
EndFunc   ;==>_SkinHotkeysCreate


; Смена языка: новые заголовок и подписи, высота прежняя
Func _SkinHotkeysSetRows($iCtrl, $sTitle, ByRef $aRows)
	Local $i = __SkinHotkeysIndexOf($iCtrl)
	If $i < 0 Then Return
	$g_aSkinHotkeys[$i][4] = $sTitle
	$g_aSkinHotkeys[$i][5] = $aRows
	__SkinHotkeysRenderTitle($i)
	__SkinHotkeysRenderColumns($i)
	__SkinHotkeysRenderBg($i)
EndFunc   ;==>_SkinHotkeysSetRows


; Позиция и размер панели. Колонки только сдвигаются вслед за фоном
Func _SkinHotkeysSetPos($iCtrl, $iX, $iY, $iW, $iH)
	Local $i = __SkinHotkeysIndexOf($iCtrl)
	If $i < 0 Then Return
	Local $bResize = ($iW <> $g_aSkinHotkeys[$i][2]) Or ($iH <> $g_aSkinHotkeys[$i][3])
	$g_aSkinHotkeys[$i][2] = $iW
	$g_aSkinHotkeys[$i][3] = $iH
	GUICtrlSetPos($iCtrl, $iX, $iY, $iW, $iH)
	If $bResize Then
		__SkinHotkeysRenderBg($i)
	Else
		__SkinHotkeysPositionColumns($i)
	EndIf
EndFunc   ;==>_SkinHotkeysSetPos


Func _SkinHotkeysSetShow($iCtrl, $bShow)
	Local $i = __SkinHotkeysIndexOf($iCtrl)
	If $i < 0 Then Return
	Local $iState = $bShow ? $GUI_SHOW : $GUI_HIDE
	GUICtrlSetState($iCtrl, $iState)
	GUICtrlSetState($g_aSkinHotkeys[$i][6], $iState)
	For $c = 0 To $gc_iSkinHkCols - 1
		GUICtrlSetState($g_aSkinHotkeys[$i][__SkinHotkeysColBase($c)], $iState)
	Next
EndFunc   ;==>_SkinHotkeysSetShow


; Докинг поменял размер панели: перерисовать фон и расставить колонки заново
Func _SkinHotkeysSyncSize()
	For $i = 0 To UBound($g_aSkinHotkeys) - 1
		Local $aSize = _SkinCtrlSize($g_aSkinHotkeys[$i][0])
		If $aSize[0] < 1 Then ContinueLoop
		If $aSize[0] = $g_aSkinHotkeys[$i][2] And $aSize[1] = $g_aSkinHotkeys[$i][3] Then ContinueLoop
		$g_aSkinHotkeys[$i][2] = $aSize[0]
		$g_aSkinHotkeys[$i][3] = $aSize[1]
		__SkinHotkeysRenderBg($i)
	Next
EndFunc   ;==>_SkinHotkeysSyncSize


Func _SkinHotkeysRenderAll()
	For $i = 0 To UBound($g_aSkinHotkeys) - 1
		__SkinHotkeysRenderTitle($i)
		__SkinHotkeysRenderColumns($i)
		__SkinHotkeysRenderBg($i)
	Next
EndFunc   ;==>_SkinHotkeysRenderAll


Func _SkinHotkeysShutdown()
	For $i = 0 To UBound($g_aSkinHotkeys) - 1
		If $g_aSkinHotkeys[$i][1] Then _WinAPI_DeleteObject($g_aSkinHotkeys[$i][1])
		If $g_aSkinHotkeys[$i][7] Then _WinAPI_DeleteObject($g_aSkinHotkeys[$i][7])
		$g_aSkinHotkeys[$i][1] = 0
		$g_aSkinHotkeys[$i][7] = 0
		For $c = 0 To $gc_iSkinHkCols - 1
			Local $iBmp = __SkinHotkeysColBase($c) + 1
			If $g_aSkinHotkeys[$i][$iBmp] Then _WinAPI_DeleteObject($g_aSkinHotkeys[$i][$iBmp])
			$g_aSkinHotkeys[$i][$iBmp] = 0
		Next
	Next

	For $sKey In MapKeys($g_mSkinHkIcons)
		If $g_mSkinHkIcons[$sKey] Then _GDIPlus_ImageDispose($g_mSkinHkIcons[$sKey])
	Next
	Local $mEmpty[]
	$g_mSkinHkIcons = $mEmpty
EndFunc   ;==>_SkinHotkeysShutdown


; ============================================================
; Внутреннее
; ============================================================

Func __SkinHotkeysIndexOf($iCtrl)
	For $i = 0 To UBound($g_aSkinHotkeys) - 1
		If $g_aSkinHotkeys[$i][0] = $iCtrl Then Return $i
	Next
	Return -1
EndFunc   ;==>__SkinHotkeysIndexOf


; Начало тройки колонки строк $iCol в реестре
Func __SkinHotkeysColBase($iCol)
	Return $gc_iSkinHkColBase + 3 * $iCol
EndFunc   ;==>__SkinHotkeysColBase


; Фон карточки – единственное, что перерисовывает ресайз. Колонки сдвигаются
; после отрисовки холста, и всё сразу выводится на экран: иначе на старом месте
; колонки до следующего WM_PAINT фона оставался её след.
Func __SkinHotkeysRenderBg($iIndex)
	Local $iW = $g_aSkinHotkeys[$iIndex][2], $iH = $g_aSkinHotkeys[$iIndex][3]
	If $iW < 1 Or $iH < 1 Then Return __SkinHotkeysPositionColumns($iIndex)

	Local $hGfx
	Local $hCanvas = _SkinCanvas($iW, $iH, _SkinArgb($g_iSkinBg), $hGfx)
	_SkinBox($hGfx, 0, 0, $iW, $iH, $gc_nSkinRadCard, _
			_SkinArgb($g_iSkinCard), _SkinArgb($g_iSkinCardBorder))
	_SkinCanvasApply($g_aSkinHotkeys[$iIndex][0], $g_aSkinHotkeys[$iIndex][1], $hCanvas, $hGfx)

	; Докинг поднимает фон в z-order, и тот закрашивал бы колонки: вниз на каждой перерисовке
	_SkinSendToBack($g_aSkinHotkeys[$iIndex][0])
	__SkinHotkeysPositionColumns($iIndex)
	__SkinHotkeysRedraw($g_aSkinHotkeys[$iIndex][0])

	; Фон уже успел закрасить колонки и заголовок, а сами они не перерисовываются
	__SkinHotkeysRedraw($g_aSkinHotkeys[$iIndex][6])
	For $c = 0 To $gc_iSkinHkCols - 1
		__SkinHotkeysRedraw($g_aSkinHotkeys[$iIndex][__SkinHotkeysColBase($c)])
	Next
EndFunc   ;==>__SkinHotkeysRenderBg


Func __SkinHotkeysRedraw($iCtrl)
	Local $hCtrl = GUICtrlGetHandle($iCtrl)
	If $hCtrl Then _WinAPI_RedrawWindow($hCtrl, 0, 0, BitOR($RDW_INVALIDATE, $RDW_UPDATENOW, $RDW_ERASE))
EndFunc   ;==>__SkinHotkeysRedraw


; Заголовок – холст по ширине текста, ресайз его не перерисовывает
Func __SkinHotkeysRenderTitle($iIndex)
	Local $sTitle = $g_aSkinHotkeys[$iIndex][4]
	Local $hFont = _SkinFont($g_nSkinSizeSmall, True)
	Local $iW = _SkinTextW($sTitle, $hFont), $iH = 16
	If $iW < 1 Then Return
	GUICtrlSetPos($g_aSkinHotkeys[$iIndex][6], Default, Default, $iW, $iH)

	Local $hGfx
	Local $hCanvas = _SkinCanvas($iW, $iH, _SkinArgb($g_iSkinCard), $hGfx)
	_SkinText($hGfx, $sTitle, 0, 0, $iW, $iH, $hFont, _SkinArgb($g_iSkinText3), 0, 1)
	_SkinCanvasApply($g_aSkinHotkeys[$iIndex][6], $g_aSkinHotkeys[$iIndex][7], $hCanvas, $hGfx)
EndFunc   ;==>__SkinHotkeysRenderTitle


; Колонки строк, каждая в холст под свой контент
Func __SkinHotkeysRenderColumns($iIndex)
	Local $aRows = $g_aSkinHotkeys[$iIndex][5]
	Local $iRows = UBound($aRows)
	Local $iPerCol = Ceiling($iRows / $gc_iSkinHkCols)
	Local $hFontSmall = _SkinFont($g_nSkinSizeSmall)
	Local $hFontKey = _SkinFont($g_nSkinSizeSmall, True)

	; Подписи колонки начинаются с общего края – от самого длинного ряда клавиш,
	; иначе левый край подписей рваный
	Local $aKeysW[$gc_iSkinHkCols], $aLabelW[$gc_iSkinHkCols]
	For $c = 0 To $gc_iSkinHkCols - 1
		$aKeysW[$c] = 0
		$aLabelW[$c] = 0
	Next
	For $i = 0 To $iRows - 1
		Local $iCol = Int($i / $iPerCol)
		Local $iKeysW = __SkinHotkeysKeysW($aRows[$i][0], $hFontKey)
		Local $iLabelW = _SkinTextW($aRows[$i][1], $hFontSmall)
		If $iKeysW > $aKeysW[$iCol] Then $aKeysW[$iCol] = $iKeysW
		If $iLabelW > $aLabelW[$iCol] Then $aLabelW[$iCol] = $iLabelW
	Next

	Local $iColH = $iPerCol * $gc_iSkinHkRowH
	For $c = 0 To $gc_iSkinHkCols - 1
		Local $iBase = __SkinHotkeysColBase($c)
		Local $iFrom = $c * $iPerCol, $iTo = _Min($iFrom + $iPerCol, $iRows)
		Local $iW = ($iFrom < $iTo) ? $aKeysW[$c] + $gc_iSkinHkKeyGap + $aLabelW[$c] : 0
		$g_aSkinHotkeys[$iIndex][$iBase + 2] = $iW
		__SkinHotkeysRenderOneColumn($g_aSkinHotkeys[$iIndex][$iBase], $g_aSkinHotkeys[$iIndex][$iBase + 1], _
				$iW, $iColH, $aRows, $iFrom, $iTo, $aKeysW[$c], $hFontSmall, $hFontKey)
	Next
EndFunc   ;==>__SkinHotkeysRenderColumns


; $iKeysW – самый длинный ряд клавиш колонки, от него общий край подписей.
; Короткий ряд прижат вправо, к подписи, иначе одиночная клавиша отрывалась от неё.
Func __SkinHotkeysRenderOneColumn($iCtrl, ByRef $hOldBmp, $iW, $iH, ByRef $aRows, $iFrom, $iTo, $iKeysW, _
		$hFontSmall, $hFontKey)
	If $iW < 1 Or $iH < 1 Then Return
	GUICtrlSetPos($iCtrl, Default, Default, $iW, $iH)

	Local $iTextX = $iKeysW + $gc_iSkinHkKeyGap
	Local $hGfx
	Local $hCanvas = _SkinCanvas($iW, $iH, _SkinArgb($g_iSkinCard), $hGfx)
	For $i = $iFrom To $iTo - 1
		Local $iRy = ($i - $iFrom) * $gc_iSkinHkRowH
		Local $iKeyX = $iKeysW - __SkinHotkeysKeysW($aRows[$i][0], $hFontKey)
		__SkinHotkeysKeys($hGfx, $aRows[$i][0], $iKeyX, $iRy, $hFontKey)
		_SkinText($hGfx, $aRows[$i][1], $iTextX, $iRy + $gc_iSkinHkTextDY, $iW - $iTextX, 17, _
				$hFontSmall, _SkinArgb($g_iSkinText2), 0, 1)
	Next
	_SkinCanvasApply($iCtrl, $hOldBmp, $hCanvas, $hGfx)
EndFunc   ;==>__SkinHotkeysRenderOneColumn


; На опорной ширине колонки разнесены от левого поля до правого с равными зазорами
; не меньше $gc_iSkinHkColGap: подряд короткие английские подписи жались влево.
; Шире опорной каждая треть карточки прибавляет поровну, и колонка едет с её
; центром: первая на 1/6 прибавки, вторая на 1/2, третья на 5/6.
Func __SkinHotkeysPositionColumns($iIndex)
	Local $aBg = ControlGetPos(_SkinGui(), "", $g_aSkinHotkeys[$iIndex][0])
	If @error Or Not IsArray($aBg) Then Return
	Local $iBgX = $aBg[0], $iBgY = $aBg[1]

	; Прибавка каждой трети. Уже опорной окно не ужимается, влево колонки не едут
	Local $nGrow = _Max(0, $g_aSkinHotkeys[$iIndex][2] - $g_aSkinHotkeys[$iIndex][8]) / $gc_iSkinHkCols

	; Зазор между колонками на опорной ширине: остаток карточки делится поровну
	Local $iCount = 0, $iSumW = 0
	For $c = 0 To $gc_iSkinHkCols - 1
		Local $iColW = $g_aSkinHotkeys[$iIndex][__SkinHotkeysColBase($c) + 2]
		If $iColW < 1 Then ContinueLoop
		$iCount += 1
		$iSumW += $iColW
	Next
	Local $nGap = $gc_iSkinHkColGap
	If $iCount > 1 Then $nGap = _Max($gc_iSkinHkColGap, _
			($g_aSkinHotkeys[$iIndex][8] - $gc_iSkinHkPad * 2 - $iSumW) / ($iCount - 1))

	Local $nColX = $iBgX + $gc_iSkinHkPad
	Local $iColY = $iBgY + $gc_iSkinHkTitleH
	GUICtrlSetPos($g_aSkinHotkeys[$iIndex][6], $iBgX + $gc_iSkinHkPad, $iBgY + 8)
	For $c = 0 To $gc_iSkinHkCols - 1
		Local $iBase = __SkinHotkeysColBase($c)
		If $g_aSkinHotkeys[$iIndex][$iBase + 2] < 1 Then ContinueLoop
		GUICtrlSetPos($g_aSkinHotkeys[$iIndex][$iBase], Round($nColX + $nGrow * ($c + 0.5)), $iColY)
		$nColX += $g_aSkinHotkeys[$iIndex][$iBase + 2] + $nGap
	Next
EndFunc   ;==>__SkinHotkeysPositionColumns


; Ширина ряда клавиш без отрисовки, по правилам __SkinHotkeysKeys
Func __SkinHotkeysKeysW($sKeys, $hFont)
	Local $aPairs = StringSplit($sKeys, "|", $STR_NOCOUNT)
	Local $iW = 0
	For $i = 0 To UBound($aPairs) - 1
		Local $aTokGlyph = StringSplit($aPairs[$i], ":", $STR_NOCOUNT)
		Local $sGlyph = (UBound($aTokGlyph) > 1) ? $aTokGlyph[1] : $aTokGlyph[0]
		$iW += (__SkinHkIcon($aTokGlyph[0]) ? $gc_iSkinHkIconSize : __SkinHkBadgeW($sGlyph, $hFont)) + 3
	Next
	Return $iW - 3
EndFunc   ;==>__SkinHotkeysKeysW


; Ширина запасного бейджа: подпись с полями, не уже 19
Func __SkinHkBadgeW($sGlyph, $hFont)
	Return _Max(19, _SkinTextW($sGlyph, $hFont) + 12)
EndFunc   ;==>__SkinHkBadgeW


; Ряд клавиш из пар "ТОКЕН:подпись": иконка клавиши, а без неё бейдж с подписью.
; Возвращает занятую ширину.
Func __SkinHotkeysKeys($hGfx, $sKeys, $iX, $iY, $hFont)
	Local $aPairs = StringSplit($sKeys, "|", $STR_NOCOUNT)
	Local $iCx = $iX
	For $i = 0 To UBound($aPairs) - 1
		Local $aTokGlyph = StringSplit($aPairs[$i], ":", $STR_NOCOUNT)
		Local $sToken = $aTokGlyph[0]
		Local $sGlyph = (UBound($aTokGlyph) > 1) ? $aTokGlyph[1] : $sToken

		Local $hIcon = __SkinHkIcon($sToken)
		If $hIcon Then
			_GDIPlus_GraphicsDrawImageRect($hGfx, $hIcon, $iCx, $iY + 1, $gc_iSkinHkIconSize, $gc_iSkinHkIconSize)
			$iCx += $gc_iSkinHkIconSize + 3
		Else
			Local $iBw = __SkinHkBadgeW($sGlyph, $hFont)
			_SkinBox($hGfx, $iCx, $iY + 1, $iBw, 16, 2, _
					_SkinArgb($g_iSkinKeyBg), _SkinArgb($g_iSkinKeyBorder))
			; Нижняя грань толще: бейдж читается клавишей, а не полем ввода
			_SkinLine($hGfx, $iCx + 3, $iY + 16, $iCx + $iBw - 3, $iY + 16, _SkinArgb($g_iSkinKeyBorder))
			_SkinText($hGfx, $sGlyph, $iCx, $iY, $iBw, 17, $hFont, _SkinArgb($g_iSkinText2), 1, 1)
			$iCx += $iBw + 3
		EndIf
	Next
	Return $iCx - 3 - $iX
EndFunc   ;==>__SkinHotkeysKeys


; Папка иконок клавиш – соседняя с папкой иконок скина
Func __SkinHkIconDir()
	Return StringRegExpReplace($g_sSkinIconDir, "\\Icons$", "\\KeyIcons")
EndFunc   ;==>__SkinHkIconDir


; Имя файла иконки по токену клавиши. Буквы и цифры – "Keyboard" & токен (KeyboardH)
Func __SkinHkIconFile($sToken)
	Switch $sToken
		Case "LEFT"
			Return "KeyboardArrowLeft"
		Case "RIGHT"
			Return "KeyboardArrowRight"
		Case "UP"
			Return "KeyboardArrowUp"
		Case "DOWN"
			Return "KeyboardArrowDown"
		Case "PLUS"
			Return "KeyboardPlus"
		Case "MINUS"
			Return "KeyboardMinus"
		Case "CTRL"
			Return "KeyboardCtrl"
		Case "ALT"
			Return "KeyboardAlt"
	EndSwitch
	Return "Keyboard" & $sToken
EndFunc   ;==>__SkinHkIconFile


; Картинка клавиши из кэша, 0 – иконки нет. Набор уже разложен по темам, не перекрашивается
Func __SkinHkIcon($sToken)
	Local $sTheme = $g_bSkinDark ? "Dark" : "Light"
	Local $sKey = $sToken & "|" & $sTheme
	If MapExists($g_mSkinHkIcons, $sKey) Then Return $g_mSkinHkIcons[$sKey]

	Local $sPath = __SkinHkIconDir() & "\" & $sTheme & "\" & __SkinHkIconFile($sToken) & ".png"
	Local $hImg = 0
	If FileExists($sPath) Then
		$hImg = _GDIPlus_ImageLoadFromFile($sPath)
		If @error Then $hImg = 0
	EndIf
	$g_mSkinHkIcons[$sKey] = $hImg
	Return $hImg
EndFunc   ;==>__SkinHkIcon
