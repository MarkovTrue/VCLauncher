#include-once
#include "SkinCore.au3"
#include <StaticConstants.au3>

; ============================================================
; SkinButton.au3 – кнопка со скруглённой рамкой
; ============================================================
; Вид рисуется в GDI+ и кладётся в Pic. Клик – штатное событие Pic: обработчик
; вешается обычным GUICtrlSetOnEvent. Состояния кнопки менять только функциями
; модуля: GUICtrlSetData на Pic грузит картинку из файла, а GUICtrlSetState
; не перерисовывает вид.
;
; Виды:
;   $SKINBTN_ICON   – иконка в рамке (выбор файла, обмен местами)
;   $SKINBTN_SUBTLE – без рамки, фон под курсором; умеет «включённое» состояние
;   $SKINBTN_TEXT   – подпись в рамке (кнопки окна настроек)
;   $SKINBTN_ACCENT – акцентная заливка; при высоте от 44 иконка встаёт над подписью

Global Const $SKINBTN_ICON = 0, $SKINBTN_SUBTLE = 1, $SKINBTN_TEXT = 2, $SKINBTN_ACCENT = 3

; Реестр кнопок. Колонки:
; 0 ControlID, 1 HBITMAP, 2 W, 3 H, 4 подпись, 5 имя иконки, 6 размер иконки,
; 7 вид, 8 включено, 9 под курсором, 10 доступна
Global $g_aSkinBtns[0][11]


Func _SkinBtnCreate($sText, $sIcon, $iIconSize, $iX, $iY, $iW, $iH, $iKind = $SKINBTN_ICON)
	Local $iCtrl = GUICtrlCreatePic("", $iX, $iY, $iW, $iH, $SS_NOTIFY)
	_SkinHandCursor($iCtrl)
	_SkinDockFixed($iCtrl)

	Local $iIndex = __SkinBtnFreeSlot()
	If $iIndex < 0 Then
		$iIndex = UBound($g_aSkinBtns)
		ReDim $g_aSkinBtns[$iIndex + 1][11]
	EndIf
	$g_aSkinBtns[$iIndex][0] = $iCtrl
	$g_aSkinBtns[$iIndex][1] = 0
	$g_aSkinBtns[$iIndex][2] = $iW
	$g_aSkinBtns[$iIndex][3] = $iH
	$g_aSkinBtns[$iIndex][4] = $sText
	$g_aSkinBtns[$iIndex][5] = $sIcon
	$g_aSkinBtns[$iIndex][6] = $iIconSize
	$g_aSkinBtns[$iIndex][7] = $iKind
	$g_aSkinBtns[$iIndex][8] = False
	$g_aSkinBtns[$iIndex][9] = False
	$g_aSkinBtns[$iIndex][10] = True

	__SkinBtnRender($iIndex)
	Return $iCtrl
EndFunc   ;==>_SkinBtnCreate


Func _SkinBtnSetText($iCtrl, $sText)
	Local $i = __SkinBtnIndexOf($iCtrl)
	If $i < 0 Or $g_aSkinBtns[$i][4] = $sText Then Return
	$g_aSkinBtns[$i][4] = $sText
	__SkinBtnRender($i)
EndFunc   ;==>_SkinBtnSetText


Func _SkinBtnSetIcon($iCtrl, $sIcon)
	Local $i = __SkinBtnIndexOf($iCtrl)
	If $i < 0 Or $g_aSkinBtns[$i][5] = $sIcon Then Return
	$g_aSkinBtns[$i][5] = $sIcon
	__SkinBtnRender($i)
EndFunc   ;==>_SkinBtnSetIcon


; «Включённое» состояние тихой кнопки: подложка акцентом (раскрытая шпаргалка)
Func _SkinBtnSetOn($iCtrl, $bOn)
	Local $i = __SkinBtnIndexOf($iCtrl)
	If $i < 0 Or $g_aSkinBtns[$i][8] = $bOn Then Return
	$g_aSkinBtns[$i][8] = $bOn
	__SkinBtnRender($i)
EndFunc   ;==>_SkinBtnSetOn


Func _SkinBtnSetEnabled($iCtrl, $bEnabled)
	Local $i = __SkinBtnIndexOf($iCtrl)
	If $i < 0 Or $g_aSkinBtns[$i][10] = $bEnabled Then Return
	$g_aSkinBtns[$i][10] = $bEnabled
	GUICtrlSetState($iCtrl, $bEnabled ? $GUI_ENABLE : $GUI_DISABLE)
	__SkinBtnRender($i)
EndFunc   ;==>_SkinBtnSetEnabled


Func _SkinBtnSetPos($iCtrl, $iX, $iY, $iW = -1, $iH = -1)
	Local $i = __SkinBtnIndexOf($iCtrl)
	If $i < 0 Then Return
	If $iW < 0 Then $iW = $g_aSkinBtns[$i][2]
	If $iH < 0 Then $iH = $g_aSkinBtns[$i][3]

	Local $bResize = ($iW <> $g_aSkinBtns[$i][2]) Or ($iH <> $g_aSkinBtns[$i][3])
	$g_aSkinBtns[$i][2] = $iW
	$g_aSkinBtns[$i][3] = $iH
	GUICtrlSetPos($iCtrl, $iX, $iY, $iW, $iH)
	If $bResize Then __SkinBtnRender($i)
EndFunc   ;==>_SkinBtnSetPos


; Удаляет кнопку. Окно, закрывающееся раньше приложения, зовёт это до GUIDelete:
; запись в реестре переживает окно, а ControlID AutoIt выдаёт заново.
; Строка помечается свободной, а не вырезается: отрисовка отдаёт элемент реестра
; в _SkinCanvasApply по ссылке, и ReDim из обработчика рвал бы её под таймером наведения.
Func _SkinBtnDelete($iCtrl)
	Local $i = __SkinBtnIndexOf($iCtrl)
	If $i < 0 Then Return
	_SkinHandCursorRemove($iCtrl)
	GUICtrlDelete($iCtrl)
	If $g_aSkinBtns[$i][1] Then _WinAPI_DeleteObject($g_aSkinBtns[$i][1])
	$g_aSkinBtns[$i][0] = 0
	$g_aSkinBtns[$i][1] = 0
EndFunc   ;==>_SkinBtnDelete


; Перерисовка всех кнопок после смены темы
Func _SkinBtnRenderAll()
	For $i = 0 To UBound($g_aSkinBtns) - 1
		__SkinBtnRender($i)
	Next
EndFunc   ;==>_SkinBtnRenderAll


Func _SkinBtnShutdown()
	For $i = 0 To UBound($g_aSkinBtns) - 1
		If Not $g_aSkinBtns[$i][0] Then ContinueLoop
		_SkinHandCursorRemove($g_aSkinBtns[$i][0])
		If $g_aSkinBtns[$i][1] Then _WinAPI_DeleteObject($g_aSkinBtns[$i][1])
		$g_aSkinBtns[$i][1] = 0
	Next
EndFunc   ;==>_SkinBtnShutdown


; Перерисовывает кнопки, которым докинг поменял размер: картинка в Pic осталась прежней
Func _SkinBtnSyncSize()
	For $i = 0 To UBound($g_aSkinBtns) - 1
		If Not $g_aSkinBtns[$i][0] Then ContinueLoop
		Local $aSize = _SkinCtrlSize($g_aSkinBtns[$i][0])
		If $aSize[0] < 1 Then ContinueLoop
		If $aSize[0] = $g_aSkinBtns[$i][2] And $aSize[1] = $g_aSkinBtns[$i][3] Then ContinueLoop
		$g_aSkinBtns[$i][2] = $aSize[0]
		$g_aSkinBtns[$i][3] = $aSize[1]
		__SkinBtnRender($i)
	Next
EndFunc   ;==>_SkinBtnSyncSize


; Опрос наведения. Pic не шлёт WM_MOUSELEAVE, поэтому положение курсора
; опрашивает общий таймер ядра (_SkinHoverRegister).
Func _SkinBtnHoverTick()
	Local $iUnder = _SkinCtrlUnderCursor()
	For $i = 0 To UBound($g_aSkinBtns) - 1
		If Not $g_aSkinBtns[$i][0] Then ContinueLoop
		Local $bHover = ($iUnder = $g_aSkinBtns[$i][0]) And $g_aSkinBtns[$i][10]
		If $bHover <> $g_aSkinBtns[$i][9] Then
			$g_aSkinBtns[$i][9] = $bHover
			__SkinBtnRender($i)
		EndIf
	Next
EndFunc   ;==>_SkinBtnHoverTick


; ============================================================
; Внутреннее
; ============================================================

Func __SkinBtnIndexOf($iCtrl)
	If Not $iCtrl Then Return -1 ; 0 стоит в свободных строках реестра
	For $i = 0 To UBound($g_aSkinBtns) - 1
		If $g_aSkinBtns[$i][0] = $iCtrl Then Return $i
	Next
	Return -1
EndFunc   ;==>__SkinBtnIndexOf


; Строка, освободившаяся после _SkinBtnDelete, или -1
Func __SkinBtnFreeSlot()
	For $i = 0 To UBound($g_aSkinBtns) - 1
		If Not $g_aSkinBtns[$i][0] Then Return $i
	Next
	Return -1
EndFunc   ;==>__SkinBtnFreeSlot


Func __SkinBtnRender($iIndex)
	If $iIndex < 0 Or $iIndex >= UBound($g_aSkinBtns) Then Return
	If Not $g_aSkinBtns[$iIndex][0] Then Return ; строка освобождена

	Local $iCtrl = $g_aSkinBtns[$iIndex][0]
	Local $iW = $g_aSkinBtns[$iIndex][2], $iH = $g_aSkinBtns[$iIndex][3]
	Local $sText = $g_aSkinBtns[$iIndex][4], $sIcon = $g_aSkinBtns[$iIndex][5]
	Local $iIconSize = $g_aSkinBtns[$iIndex][6], $iKind = $g_aSkinBtns[$iIndex][7]
	Local $bOn = $g_aSkinBtns[$iIndex][8], $bHover = $g_aSkinBtns[$iIndex][9]
	Local $bEnabled = $g_aSkinBtns[$iIndex][10]
	If $iW < 1 Or $iH < 1 Then Return

	Local $hGfx
	; Pic прямоугольный, а рамка скруглена: под углами должен быть фон окна
	Local $hCanvas = _SkinCanvas($iW, $iH, _SkinArgb($g_iSkinBg), $hGfx)

	Local $iFg = $bEnabled ? $g_iSkinText2 : $g_iSkinText3
	Local $nRad = $gc_nSkinRadCtrl

	Switch $iKind
		Case $SKINBTN_SUBTLE
			If $bOn Then
				_SkinBox($hGfx, 0, 0, $iW, $iH, $nRad, _
						_SkinArgb($g_iSkinAccent, 34), _SkinArgb($g_iSkinAccent, 110))
				$iFg = $g_iSkinAccent
			ElseIf $bHover Then
				_SkinFill($hGfx, 0, 0, $iW, $iH, $nRad, _SkinArgb($g_iSkinHover))
			EndIf

		Case $SKINBTN_ACCENT
			Local $iFill = $bHover ? $g_iSkinAccentHot : $g_iSkinAccent
			If Not $bEnabled Then $iFill = $g_iSkinTrack
			_SkinBox($hGfx, 0, 0, $iW, $iH, $nRad, _SkinArgb($iFill), _
					_SkinArgb($bEnabled ? $g_iSkinAccentEdge : $g_iSkinTrackBorder))
			$iFg = $bEnabled ? $g_iSkinOnAccent : $g_iSkinText3

		Case Else ; $SKINBTN_ICON, $SKINBTN_TEXT
			; Рамка ровная по периметру: подчёркивание снизу делало нижние углы площе
			Local $iBg = $bHover ? $g_iSkinHover : $g_iSkinCtrlBg
			_SkinBox($hGfx, 0, 0, $iW, $iH, $nRad, _SkinArgb($iBg), _SkinArgb($g_iSkinCtrlBorder))
	EndSwitch

	__SkinBtnContent($hGfx, $iW, $iH, $sText, $sIcon, $iIconSize, $iKind, $iFg)
	_SkinCanvasApply($iCtrl, $g_aSkinBtns[$iIndex][1], $hCanvas, $hGfx)
EndFunc   ;==>__SkinBtnRender


; Иконка и подпись внутри кнопки. В высокой кнопке иконка над подписью,
; в остальных – в строку по центру.
Func __SkinBtnContent($hGfx, $iW, $iH, $sText, $sIcon, $iIconSize, $iKind, $iFg)
	Local $bSemi = ($iKind = $SKINBTN_ACCENT)
	Local $hFont = _SkinFont(0, $bSemi)

	If $sText = "" Then
		If $sIcon <> "" Then _SkinDrawIcon($hGfx, $sIcon, $iIconSize, $iFg, _
				Int(($iW - $iIconSize) / 2), Int(($iH - $iIconSize) / 2))
		Return
	EndIf

	Local $iTextW = _SkinTextW($sText, $hFont)

	If $sIcon = "" Then
		_SkinText($hGfx, $sText, 0, 0, $iW, $iH, $hFont, _SkinArgb($iFg), 1, 1)
		Return
	EndIf

	If $iH >= 44 Then ; иконка сверху, подпись снизу
		_SkinDrawIcon($hGfx, $sIcon, $iIconSize, $iFg, Int(($iW - $iIconSize) / 2), Int(($iH - $iIconSize - 20) / 2))
		_SkinText($hGfx, $sText, 0, $iH - 24, $iW, 20, $hFont, _SkinArgb($iFg), 1, 1)
		Return
	EndIf

	Local $iBlockW = $iIconSize + 6 + $iTextW
	Local $iX = Int(($iW - $iBlockW) / 2)
	_SkinDrawIcon($hGfx, $sIcon, $iIconSize, $iFg, $iX, Int(($iH - $iIconSize) / 2))
	_SkinText($hGfx, $sText, $iX + $iIconSize + 6, 0, $iTextW + 4, $iH, $hFont, _SkinArgb($iFg), 0, 1)
EndFunc   ;==>__SkinBtnContent
