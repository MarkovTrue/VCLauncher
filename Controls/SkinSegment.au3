#include-once
#include "SkinCore.au3"
#include <StaticConstants.au3>
#include <StringConstants.au3>

; ============================================================
; SkinSegment.au3 – переключатель-сегменты вместо радиокнопок
; ============================================================
; Группа рисуется в GDI+ одной картинкой в Pic: дорожка, выбранный сегмент
; пилюлей, иконка и подпись. Клик разбирается по X курсора.
; Ширина сегмента – по его собственной подписи: по самой длинной короткие
; превращались в пустые полосы. _SegSetWidth растягивает группу до соседних
; строк, прибавка уходит сначала узким сегментам.

; Метрика: поле дорожки, поля внутри сегмента, зазор иконка-подпись, размер иконки
Global Const $gc_iSegPad = 3, $gc_iSegPadX = 10, $gc_iSegGap = 6, $gc_iSegIcon = 14

; Реестр групп. Колонки:
; 0 ControlID Pic, 1 HBITMAP, 2 X, 3 Y, 4 W, 5 H, 6 подписи "A|B",
; 7 иконки "A|B", 8 выбранный, 9 подсвеченный (-1 нет), 10 ширины "w1|w2",
; 11 имя функции-обработчика смены, 12 заданная ширина группы (0 – по подписям),
; 13 ширины по подписям "w1|w2" – от них иконка с подписью встаёт по центру сегмента,
; 14 подсказки "A|B" ("" – без подсказок), 15 функция установки подсказки,
; 16 сегмент, чья подсказка сейчас стоит на Pic (-1 никакой)
Global $g_aSegCtrls[0][17]


; Создаёт группу и возвращает ControlID её Pic. $sIcons – имена PNG через "|".
; Пустое имя (или "" на всю группу) – сегмент из одной подписи, без места под иконку.
Func _SegCreate($sTexts, $sIcons, $iX, $iY, $iH = 28, $iSelected = 0, $sOnChange = "")
	Local $sWidths, $sNatural
	Local $iW = __SegLayout($sTexts, $sIcons, 0, $sWidths, $sNatural)

	Local $iCtrl = GUICtrlCreatePic("", $iX, $iY, $iW, $iH, $SS_NOTIFY)
	GUICtrlSetOnEvent($iCtrl, "_SegOnClick")
	_SkinHandCursor($iCtrl)
	_SkinDockFixed($iCtrl)

	Local $iIndex = __SegFreeSlot()
	If $iIndex < 0 Then
		$iIndex = UBound($g_aSegCtrls)
		ReDim $g_aSegCtrls[$iIndex + 1][17]
	EndIf
	$g_aSegCtrls[$iIndex][0] = $iCtrl
	$g_aSegCtrls[$iIndex][1] = 0
	$g_aSegCtrls[$iIndex][2] = $iX
	$g_aSegCtrls[$iIndex][3] = $iY
	$g_aSegCtrls[$iIndex][4] = $iW
	$g_aSegCtrls[$iIndex][5] = $iH
	$g_aSegCtrls[$iIndex][6] = $sTexts
	$g_aSegCtrls[$iIndex][7] = $sIcons
	$g_aSegCtrls[$iIndex][8] = $iSelected
	$g_aSegCtrls[$iIndex][9] = -1
	$g_aSegCtrls[$iIndex][10] = $sWidths
	$g_aSegCtrls[$iIndex][11] = $sOnChange
	$g_aSegCtrls[$iIndex][12] = 0
	$g_aSegCtrls[$iIndex][13] = $sNatural
	$g_aSegCtrls[$iIndex][14] = ""
	$g_aSegCtrls[$iIndex][15] = ""
	$g_aSegCtrls[$iIndex][16] = -1

	__SegRender($iIndex)
	Return $iCtrl
EndFunc   ;==>_SegCreate


; Подсказки сегментов, тексты через "|". $sSetTip – функция окна (ControlID, текст),
; которая ставит подсказку по-своему; пустая – штатный GUICtrlSetTip.
Func _SegSetTips($iCtrl, $sTips, $sSetTip = "")
	Local $i = __SegIndexOf($iCtrl)
	If $i < 0 Then Return
	$g_aSegCtrls[$i][14] = $sTips
	$g_aSegCtrls[$i][15] = $sSetTip
	$g_aSegCtrls[$i][16] = -1 ; тексты могли смениться – поставить заново
	__SegApplyTip($i, ($g_aSegCtrls[$i][9] >= 0) ? $g_aSegCtrls[$i][9] : $g_aSegCtrls[$i][8])
EndFunc   ;==>_SegSetTips


; Растягивает группу до $iW, чтобы её край совпал с соседними строками.
; Уже своих подписей группа не становится. 0 – снова ширина по подписям.
Func _SegSetWidth($iCtrl, $iW)
	Local $i = __SegIndexOf($iCtrl)
	If $i < 0 Then Return
	$g_aSegCtrls[$i][12] = $iW
	__SegApplyLayout($i)
EndFunc   ;==>_SegSetWidth


; Ширина группы: вызывающему для раскладки соседних контролов
Func _SegWidth($iCtrl)
	Local $i = __SegIndexOf($iCtrl)
	Return ($i < 0) ? 0 : $g_aSegCtrls[$i][4]
EndFunc   ;==>_SegWidth


Func _SegGetSel($iCtrl)
	Local $i = __SegIndexOf($iCtrl)
	Return ($i < 0) ? -1 : $g_aSegCtrls[$i][8]
EndFunc   ;==>_SegGetSel


; Новые подписи (смена языка): ширины пересчитываются, контрол меняет размер
Func _SegSetTexts($iCtrl, $sTexts)
	Local $i = __SegIndexOf($iCtrl)
	If $i < 0 Then Return
	$g_aSegCtrls[$i][6] = $sTexts
	__SegApplyLayout($i)
EndFunc   ;==>_SegSetTexts


Func _SegSetPos($iCtrl, $iX, $iY)
	Local $i = __SegIndexOf($iCtrl)
	If $i < 0 Then Return
	$g_aSegCtrls[$i][2] = $iX
	$g_aSegCtrls[$i][3] = $iY
	GUICtrlSetPos($iCtrl, $iX, $iY, $g_aSegCtrls[$i][4], $g_aSegCtrls[$i][5])
EndFunc   ;==>_SegSetPos


; Удаляет группу. Окно, закрывающееся раньше приложения, зовёт это до GUIDelete:
; запись в реестре переживает окно, а ControlID AutoIt выдаёт заново.
; Строка помечается свободной, а не вырезается: см. _SkinBtnDelete.
Func _SegDelete($iCtrl)
	Local $i = __SegIndexOf($iCtrl)
	If $i < 0 Then Return
	_SkinHandCursorRemove($iCtrl)
	GUICtrlDelete($iCtrl)
	If $g_aSegCtrls[$i][1] Then _WinAPI_DeleteObject($g_aSegCtrls[$i][1])
	$g_aSegCtrls[$i][0] = 0
	$g_aSegCtrls[$i][1] = 0
EndFunc   ;==>_SegDelete


Func _SegRenderAll()
	For $i = 0 To UBound($g_aSegCtrls) - 1
		__SegRender($i)
	Next
EndFunc   ;==>_SegRenderAll


Func _SegShutdown()
	For $i = 0 To UBound($g_aSegCtrls) - 1
		If Not $g_aSegCtrls[$i][0] Then ContinueLoop
		_SkinHandCursorRemove($g_aSegCtrls[$i][0])
		If $g_aSegCtrls[$i][1] Then _WinAPI_DeleteObject($g_aSegCtrls[$i][1])
		$g_aSegCtrls[$i][1] = 0
	Next
EndFunc   ;==>_SegShutdown


; Клик по группе: сегмент определяется X курсора
Func _SegOnClick()
	Local $i = __SegIndexOf(@GUI_CtrlId)
	If $i < 0 Then Return

	Local $iHit = __SegHitTest($i)
	If $iHit < 0 Or $iHit = $g_aSegCtrls[$i][8] Then Return

	$g_aSegCtrls[$i][8] = $iHit
	__SegRender($i)
	If $g_aSegCtrls[$i][11] <> "" Then Call($g_aSegCtrls[$i][11], $g_aSegCtrls[$i][0])
EndFunc   ;==>_SegOnClick


; Подсветка сегмента под курсором. Pic не шлёт WM_MOUSELEAVE, поэтому опрос таймером ядра
Func _SegHoverTick()
	Local $iUnder = _SkinCtrlUnderCursor()
	For $i = 0 To UBound($g_aSegCtrls) - 1
		If Not $g_aSegCtrls[$i][0] Then ContinueLoop
		Local $iHover = -1
		If $iUnder = $g_aSegCtrls[$i][0] Then $iHover = __SegHitTest($i)
		If $iHover <> $g_aSegCtrls[$i][9] Then
			$g_aSegCtrls[$i][9] = $iHover
			__SegRender($i)
			If $iHover >= 0 Then __SegApplyTip($i, $iHover)
		EndIf
	Next
EndFunc   ;==>_SegHoverTick


; ============================================================
; Внутреннее
; ============================================================

Func __SegIndexOf($iCtrl)
	If Not $iCtrl Then Return -1 ; 0 стоит в свободных строках реестра
	For $i = 0 To UBound($g_aSegCtrls) - 1
		If $g_aSegCtrls[$i][0] = $iCtrl Then Return $i
	Next
	Return -1
EndFunc   ;==>__SegIndexOf


; Подсказка сегмента $iSeg на общем Pic группы. Ставится только при смене сегмента:
; GUICtrlSetTip пересоздаёт окно подсказки.
Func __SegApplyTip($iIndex, $iSeg)
	If $g_aSegCtrls[$iIndex][14] = "" Or $iSeg < 0 Or $iSeg = $g_aSegCtrls[$iIndex][16] Then Return

	Local $aTips = StringSplit($g_aSegCtrls[$iIndex][14], "|", $STR_NOCOUNT)
	Local $sTip = ($iSeg < UBound($aTips)) ? $aTips[$iSeg] : ""
	$g_aSegCtrls[$iIndex][16] = $iSeg
	If $g_aSegCtrls[$iIndex][15] <> "" Then
		Call($g_aSegCtrls[$iIndex][15], $g_aSegCtrls[$iIndex][0], $sTip)
	Else
		GUICtrlSetTip($g_aSegCtrls[$iIndex][0], $sTip)
	EndIf
EndFunc   ;==>__SegApplyTip


; Строка, освободившаяся после _SegDelete, или -1
Func __SegFreeSlot()
	For $i = 0 To UBound($g_aSegCtrls) - 1
		If Not $g_aSegCtrls[$i][0] Then Return $i
	Next
	Return -1
EndFunc   ;==>__SegFreeSlot


; Индекс сегмента под курсором, -1 если курсор вне группы
Func __SegHitTest($iIndex)
	Local $iLocalX = _SkinCursorLocalX($g_aSegCtrls[$iIndex][0]) - $gc_iSegPad
	If $iLocalX < 0 Then Return -1

	Local $aWidths = StringSplit($g_aSegCtrls[$iIndex][10], "|", $STR_NOCOUNT)
	Local $iEdge = 0
	For $i = 0 To UBound($aWidths) - 1
		$iEdge += Int($aWidths[$i])
		If $iLocalX < $iEdge Then Return $i
	Next
	Return -1
EndFunc   ;==>__SegHitTest


; Пересчёт ширин из подписей и заданной ширины группы, размер и перерисовка
Func __SegApplyLayout($iIndex)
	Local $sWidths, $sNatural
	Local $iW = __SegLayout($g_aSegCtrls[$iIndex][6], $g_aSegCtrls[$iIndex][7], _
			$g_aSegCtrls[$iIndex][12], $sWidths, $sNatural)

	$g_aSegCtrls[$iIndex][4] = $iW
	$g_aSegCtrls[$iIndex][10] = $sWidths
	$g_aSegCtrls[$iIndex][13] = $sNatural
	GUICtrlSetPos($g_aSegCtrls[$iIndex][0], $g_aSegCtrls[$iIndex][2], $g_aSegCtrls[$iIndex][3], _
			$iW, $g_aSegCtrls[$iIndex][5])
	__SegRender($iIndex)
EndFunc   ;==>__SegApplyLayout


; Ширины сегментов "w1|w2": по подписям ($sNatural) и растянутые до $iTargetW ($sWidths).
; Возвращает ширину группы. Уже подписей группа не становится.
Func __SegLayout($sTexts, $sIcons, $iTargetW, ByRef $sWidths, ByRef $sNatural)
	Local $aWidths = __SegMeasure($sTexts, $sIcons)
	Local $iSum = 0
	$sNatural = ""
	For $i = 0 To UBound($aWidths) - 1
		$iSum += $aWidths[$i]
		$sNatural &= (($i = 0) ? "" : "|") & $aWidths[$i]
	Next

	If $iTargetW - $gc_iSegPad * 2 > $iSum Then __SegStretch($aWidths, $iTargetW - $gc_iSegPad * 2)

	Local $iW = $gc_iSegPad * 2
	$sWidths = ""
	For $i = 0 To UBound($aWidths) - 1
		$iW += $aWidths[$i]
		$sWidths &= (($i = 0) ? "" : "|") & $aWidths[$i]
	Next
	Return $iW
EndFunc   ;==>__SegLayout


; Растягивает сегменты до суммы $iTotal. Прибавка достаётся сначала узким:
; группа стремится к равным долям, а сегмент шире равной доли остаётся как есть.
Func __SegStretch(ByRef $aWidths, $iTotal)
	Local $iN = UBound($aWidths)
	Local $aKeep[$iN]
	Local $iFree = $iTotal, $iCount = $iN

	Local $bAgain = True
	While $bAgain
		$bAgain = False
		For $i = 0 To $iN - 1
			If $aKeep[$i] Or $aWidths[$i] * $iCount <= $iFree Then ContinueLoop
			$aKeep[$i] = True
			$iFree -= $aWidths[$i]
			$iCount -= 1
			$bAgain = True
		Next
	WEnd

	; Остаток поровну, пиксели от деления – последнему растянутому
	For $i = 0 To $iN - 1
		If $aKeep[$i] Then ContinueLoop
		$aWidths[$i] = ($iCount = 1) ? $iFree : Int($iFree / $iCount)
		$iFree -= $aWidths[$i]
		$iCount -= 1
	Next
EndFunc   ;==>__SegStretch


; Ширина сегмента: поля, иконка с зазором и своя подпись. Без иконки место под неё
; не резервируется, иначе подпись уезжает вправо.
Func __SegMeasure($sTexts, $sIcons)
	Local $aTexts = StringSplit($sTexts, "|", $STR_NOCOUNT)
	Local $aIcons = StringSplit($sIcons, "|", $STR_NOCOUNT)
	Local $aWidths[UBound($aTexts)]
	Local $hFont = _SkinFont()

	For $i = 0 To UBound($aTexts) - 1
		$aWidths[$i] = $gc_iSegPadX * 2 + _SkinTextW($aTexts[$i], $hFont)
		If __SegIcon($aIcons, $i) <> "" Then $aWidths[$i] += $gc_iSegIcon + $gc_iSegGap
	Next
	Return $aWidths
EndFunc   ;==>__SegMeasure


; Имя иконки сегмента без расширения. "" – сегмент из одной подписи
Func __SegIcon(ByRef $aIcons, $iIndex)
	If $iIndex >= UBound($aIcons) Then Return ""
	Return StringReplace($aIcons[$iIndex], ".png", "")
EndFunc   ;==>__SegIcon


Func __SegRender($iIndex)
	If $iIndex < 0 Or $iIndex >= UBound($g_aSegCtrls) Then Return
	If Not $g_aSegCtrls[$iIndex][0] Then Return ; строка освобождена

	Local $iW = $g_aSegCtrls[$iIndex][4], $iH = $g_aSegCtrls[$iIndex][5]
	If $iW < 1 Or $iH < 1 Then Return

	Local $aTexts = StringSplit($g_aSegCtrls[$iIndex][6], "|", $STR_NOCOUNT)
	Local $aIcons = StringSplit($g_aSegCtrls[$iIndex][7], "|", $STR_NOCOUNT)
	Local $aWidths = StringSplit($g_aSegCtrls[$iIndex][10], "|", $STR_NOCOUNT)
	Local $aNatural = StringSplit($g_aSegCtrls[$iIndex][13], "|", $STR_NOCOUNT)
	Local $iSel = $g_aSegCtrls[$iIndex][8], $iHover = $g_aSegCtrls[$iIndex][9]

	Local $hGfx
	Local $hCanvas = _SkinCanvas($iW, $iH, _SkinArgb($g_iSkinBg), $hGfx)

	; Дорожка
	_SkinBox($hGfx, 0, 0, $iW, $iH, $gc_nSkinRadTrack, _
			_SkinArgb($g_iSkinTrack), _SkinArgb($g_iSkinTrackBorder))

	Local $hFont = _SkinFont()
	Local $hFontSel = _SkinFont(0, True)
	Local $iX = $gc_iSegPad

	For $i = 0 To UBound($aTexts) - 1
		Local $iSegW = Int($aWidths[$i])
		Local $bOn = ($i = $iSel)
		Local $iFg = $g_iSkinText2, $iIconFg = $g_iSkinText3

		If $bOn Then
			; Пилюля без тени снизу: смещённая копия утяжеляла нижние углы
			_SkinBox($hGfx, $iX, $gc_iSegPad, $iSegW, $iH - $gc_iSegPad * 2, $gc_nSkinRadCtrl, _
					_SkinArgb($g_iSkinSegSelBg), _SkinArgb($g_iSkinSegSelBorder))
			$iFg = $g_iSkinSegSelText
			$iIconFg = $iFg ; иконка выбранного сегмента цветом текста, не акцентом
		ElseIf $i = $iHover Then
			_SkinFill($hGfx, $iX, $gc_iSegPad, $iSegW, $iH - $gc_iSegPad * 2, _
					$gc_nSkinRadCtrl, _SkinArgb($g_iSkinHover))
			$iFg = $g_iSkinText1
		EndIf

		Local $hTextFont = $bOn ? $hFontSel : $hFont
		Local $sIcon = __SegIcon($aIcons, $i)
		If $sIcon = "" Then
			; Без иконки подпись стоит по центру всей ширины сегмента
			_SkinText($hGfx, $aTexts[$i], $iX, 0, $iSegW, $iH, $hTextFont, _SkinArgb($iFg), 1, 1)
		Else
			; Иконка с подписью – одним блоком, в растянутом сегменте тоже по центру
			Local $iIconX = $iX + $gc_iSegPadX
			If $i < UBound($aNatural) Then $iIconX += Int(($iSegW - Int($aNatural[$i])) / 2)
			_SkinDrawIcon($hGfx, $sIcon, $gc_iSegIcon, $iIconFg, $iIconX, Int(($iH - $gc_iSegIcon) / 2))

			Local $iTextX = $iIconX + $gc_iSegIcon + $gc_iSegGap
			_SkinText($hGfx, $aTexts[$i], $iTextX, 0, $iX + $iSegW - $iTextX, $iH, _
					$hTextFont, _SkinArgb($iFg), 0, 1)
		EndIf

		$iX += $iSegW
	Next

	_SkinCanvasApply($g_aSegCtrls[$iIndex][0], $g_aSegCtrls[$iIndex][1], $hCanvas, $hGfx)
EndFunc   ;==>__SegRender
