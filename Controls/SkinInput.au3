#include-once
#include "SkinCore.au3"
#include <EditConstants.au3>
#include <StaticConstants.au3>
#include <WindowsConstants.au3>

; ============================================================
; SkinInput.au3 – поле ввода со скруглённой рамкой
; ============================================================
; Рамка рисуется в GDI+ и кладётся в Pic, поверх лежит Edit без рамки с той же
; заливкой. Возвращается ControlID именно Edit: приложение работает с полем
; обычными GUICtrlSetData и GUICtrlRead.
;
; Edit любого вида рисует первую строку от верха клиентской области, поэтому
; по центру рамки ставится строчный бокс высотой _SkinGdiLineHeight, а не контрол.
; Видимый отступ текста от рамки – $gc_iSkinInputPadX плюс собственные 3 px
; EM_GETMARGINS, один на оба вида.

Global Const $gc_iSkinInputPadX = 7  ; поле от рамки до клиентской области Edit

; Реестр полей. Колонки:
; 0 ControlID Edit, 1 ControlID Pic-рамки, 2 HBITMAP, 3 W, 4 H,
; 5 в фокусе, 6 под курсором, 7 многострочное, 8 X, 9 Y
Global $g_aSkinInputs[0][10]


; Создаёт поле. $iStyle – дополнительные стили Edit ($ES_RIGHT и т.п.).
; $bMultiline – многострочное поле (команда): Edit занимает рамку целиком.
Func _SkinInputCreate($iX, $iY, $iW, $iH, $iStyle = -1, $bMultiline = False)
	; $WS_CLIPSIBLINGS обязателен: иначе рамка при перерисовке закрашивает лежащий на ней Edit
	Local $iPic = GUICtrlCreatePic("", $iX, $iY, $iW, $iH, $WS_CLIPSIBLINGS)

	; Без полосы прокрутки: штатная выглядит заплаткой внутри скруглённой рамки.
	; $ES_AUTOVSCROLL оставляет прокрутку колесом и кареткой.
	Local $iEditStyle = $bMultiline _
			? BitOR($ES_LEFT, $ES_MULTILINE, $ES_AUTOVSCROLL) _
			: BitOR($ES_LEFT, $ES_AUTOHSCROLL)
	If $iStyle <> -1 Then $iEditStyle = BitOR($iEditStyle, $iStyle)

	; Последний аргумент 0 снимает $WS_EX_CLIENTEDGE: рамку рисуем сами
	Local $aRect = __SkinInputEditRect($iX, $iY, $iW, $iH, $bMultiline)
	Local $iEdit = GUICtrlCreateInput("", $aRect[0], $aRect[1], $aRect[2], $aRect[3], $iEditStyle, 0)
	_SkinDockFixed($iPic)
	_SkinDockFixed($iEdit)
	; Рамка – подложка под Edit, иначе она перехватывает у него клики
	_SkinSendToBack($iPic)

	Local $iIndex = UBound($g_aSkinInputs)
	ReDim $g_aSkinInputs[$iIndex + 1][10]
	$g_aSkinInputs[$iIndex][0] = $iEdit
	$g_aSkinInputs[$iIndex][1] = $iPic
	$g_aSkinInputs[$iIndex][2] = 0
	$g_aSkinInputs[$iIndex][3] = $iW
	$g_aSkinInputs[$iIndex][4] = $iH
	$g_aSkinInputs[$iIndex][5] = False
	$g_aSkinInputs[$iIndex][6] = False
	$g_aSkinInputs[$iIndex][7] = $bMultiline
	$g_aSkinInputs[$iIndex][8] = $iX
	$g_aSkinInputs[$iIndex][9] = $iY

	__SkinInputRender($iIndex)
	Return $iEdit
EndFunc   ;==>_SkinInputCreate


; Pic-рамка поля: вызывающему для докинга
Func _SkinInputFrame($iEdit)
	Local $i = __SkinInputIndexOf($iEdit)
	Return ($i < 0) ? 0 : $g_aSkinInputs[$i][1]
EndFunc   ;==>_SkinInputFrame


Func _SkinInputSetPos($iEdit, $iX, $iY, $iW = -1, $iH = -1)
	Local $i = __SkinInputIndexOf($iEdit)
	If $i < 0 Then Return
	If $iW < 0 Then $iW = $g_aSkinInputs[$i][3]
	If $iH < 0 Then $iH = $g_aSkinInputs[$i][4]

	Local $bResize = ($iW <> $g_aSkinInputs[$i][3]) Or ($iH <> $g_aSkinInputs[$i][4])
	$g_aSkinInputs[$i][3] = $iW
	$g_aSkinInputs[$i][4] = $iH
	$g_aSkinInputs[$i][8] = $iX
	$g_aSkinInputs[$i][9] = $iY

	GUICtrlSetPos($g_aSkinInputs[$i][1], $iX, $iY, $iW, $iH)
	Local $aRect = __SkinInputEditRect($iX, $iY, $iW, $iH, $g_aSkinInputs[$i][7])
	GUICtrlSetPos($iEdit, $aRect[0], $aRect[1], $aRect[2], $aRect[3])
	If $bResize Then __SkinInputRender($i)
EndFunc   ;==>_SkinInputSetPos


Func _SkinInputRenderAll()
	For $i = 0 To UBound($g_aSkinInputs) - 1
		__SkinInputRender($i)
	Next
EndFunc   ;==>_SkinInputRenderAll


Func _SkinInputShutdown()
	For $i = 0 To UBound($g_aSkinInputs) - 1
		If $g_aSkinInputs[$i][2] Then _WinAPI_DeleteObject($g_aSkinInputs[$i][2])
		$g_aSkinInputs[$i][2] = 0
	Next
EndFunc   ;==>_SkinInputShutdown


; Перерисовывает рамки, которым докинг поменял размер, и заново кладёт Edit
; внутрь: высоту докинг отдаёт любую, а многострочному полю нужна кратная строке.
; X и Y сохранённые: все поля привязаны левым верхним углом.
Func _SkinInputSyncSize()
	For $i = 0 To UBound($g_aSkinInputs) - 1
		Local $aSize = _SkinCtrlSize($g_aSkinInputs[$i][1])
		If $aSize[0] < 1 Then ContinueLoop
		If $aSize[0] = $g_aSkinInputs[$i][3] And $aSize[1] = $g_aSkinInputs[$i][4] Then ContinueLoop
		$g_aSkinInputs[$i][3] = $aSize[0]
		$g_aSkinInputs[$i][4] = $aSize[1]

		Local $aRect = __SkinInputEditRect($g_aSkinInputs[$i][8], $g_aSkinInputs[$i][9], _
				$aSize[0], $aSize[1], $g_aSkinInputs[$i][7])
		GUICtrlSetPos($g_aSkinInputs[$i][0], $aRect[0], $aRect[1], $aRect[2], $aRect[3])
		__SkinInputRender($i)
	Next
EndFunc   ;==>_SkinInputSyncSize


; Опрос фокуса и наведения. Фокус читается, а не ловится EN_SETFOCUS:
; модуль не зависит от обработчика WM_COMMAND приложения.
Func _SkinInputHoverTick()
	Local $iUnder = _SkinCtrlUnderCursor()
	Local $hFocus = _WinAPI_GetFocus()

	For $i = 0 To UBound($g_aSkinInputs) - 1
		Local $iEdit = $g_aSkinInputs[$i][0]
		Local $bFocus = ($hFocus = GUICtrlGetHandle($iEdit))
		Local $bHover = ($iUnder = $iEdit) Or ($iUnder = $g_aSkinInputs[$i][1])
		If $bFocus <> $g_aSkinInputs[$i][5] Or $bHover <> $g_aSkinInputs[$i][6] Then
			$g_aSkinInputs[$i][5] = $bFocus
			$g_aSkinInputs[$i][6] = $bHover
			__SkinInputRender($i)
		EndIf
	Next
EndFunc   ;==>_SkinInputHoverTick


; ============================================================
; Внутреннее
; ============================================================

Func __SkinInputIndexOf($iEdit)
	For $i = 0 To UBound($g_aSkinInputs) - 1
		If $g_aSkinInputs[$i][0] = $iEdit Then Return $i
	Next
	Return -1
EndFunc   ;==>__SkinInputIndexOf


; Прямоугольник Edit внутри рамки: [x, y, w, h]. По вертикали центрируется строчный бокс
Func __SkinInputEditRect($iX, $iY, $iW, $iH, $bMultiline)
	Local $aRect[4]
	Local $iLineH = _SkinGdiLineHeight()

	Local $iTextH
	If $bMultiline Then
		; Высота – целое число строк, иначе последняя режется по горизонтали
		$iTextH = Int(($iH - 12) / $iLineH) * $iLineH
		If $iTextH < $iLineH Then $iTextH = $iLineH
	Else
		$iTextH = $iLineH
	EndIf

	Local $iTop = Int(($iH - $iTextH) / 2)
	If $iTop < 0 Then $iTop = 0

	; Однострочному вся высота до нижней грани: клик по пустому низу попадает в поле.
	; Саму грань не перекрываем: непрозрачный Edit стёр бы скруглённый угол.
	Local $iEditH = $iTextH
	If Not $bMultiline Then
		$iEditH = $iH - $iTop - 1
		If $iEditH < $iTextH Then $iEditH = $iTextH
	EndIf

	$aRect[0] = $iX + $gc_iSkinInputPadX
	$aRect[1] = $iY + $iTop
	$aRect[2] = $iW - $gc_iSkinInputPadX * 2
	$aRect[3] = $iEditH
	Return $aRect
EndFunc   ;==>__SkinInputEditRect


Func __SkinInputRender($iIndex)
	Local $iW = $g_aSkinInputs[$iIndex][3], $iH = $g_aSkinInputs[$iIndex][4]
	If $iW < 1 Or $iH < 1 Then Return

	Local $hGfx
	; Pic прямоугольный, рамка скруглена: под углами должен лежать фон окна
	Local $hCanvas = _SkinCanvas($iW, $iH, _SkinArgb($g_iSkinBg), $hGfx)

	; Фокус и наведение – цветом рамки: полоска снизу делала нижние углы площе верхних
	Local $bFocus = $g_aSkinInputs[$iIndex][5], $bHover = $g_aSkinInputs[$iIndex][6]
	Local $iBorder = $g_iSkinCtrlBorder
	If $bHover Then $iBorder = $g_iSkinCtrlUnder
	If $bFocus Then $iBorder = $g_iSkinAccent
	_SkinBox($hGfx, 0, 0, $iW, $iH, $gc_nSkinRadCtrl, _SkinArgb($g_iSkinCtrlBg), _SkinArgb($iBorder))

	_SkinCanvasApply($g_aSkinInputs[$iIndex][1], $g_aSkinInputs[$iIndex][2], $hCanvas, $hGfx)

	; Рамка при перерисовке успевает мазнуть по Edit, и без его перерисовки
	; текст пропадал до следующего ввода
	Local $hEdit = GUICtrlGetHandle($g_aSkinInputs[$iIndex][0])
	If $hEdit Then _WinAPI_RedrawWindow($hEdit, 0, 0, BitOR($RDW_INVALIDATE, $RDW_UPDATENOW))
EndFunc   ;==>__SkinInputRender
