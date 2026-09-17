#include-once
#include "SkinCore.au3"
#include <StaticConstants.au3>
#include <WindowsConstants.au3>

; ============================================================
; SkinHeader.au3 – шапка окна: логотип, название, версия
; ============================================================
; Кнопки шапки вызывающий создаёт поверх этого Pic обычными SkinButton.
; Картинка размером под свой контент и прибита к левому краю: справа тот же фон
; окна, шва не видно, а ресайз шапку не перерисовывает. Полноширинная
; перерисовка на каждом шаге перетаскивания рамки заметно тормозила.

Global Const $gc_iSkinHeaderH = 60   ; высота строки представления
Global Const $gc_iSkinHeaderLogo = 32
Global Const $gc_iSkinHeaderPadL = 16, $gc_iSkinHeaderTextGap = 12, $gc_iSkinHeaderPadR = 16
Global Const $gc_iSkinHeaderMaxW = 400 ; потолок: длинный заголовок не наедет на кнопки шапки

; Реестр шапок. Колонки:
; 0 ControlID Pic, 1 HBITMAP, 2 W, 3 H, 4 путь к логотипу, 5 название, 6 подпись
Global $g_aSkinHeaders[0][7]


Func _SkinHeaderCreate($iX, $iY, $sLogoPath, $sTitle, $sSubtitle, $iH = $gc_iSkinHeaderH)
	Local $iW = __SkinHeaderContentW($sTitle, $sSubtitle)

	; $WS_CLIPSIBLINGS обязателен: без него перерисовка шапки закрашивает лежащие
	; на ней кнопки, и они пропадали при старте и при смене темы
	Local $iCtrl = GUICtrlCreatePic("", $iX, $iY, $iW, $iH, $WS_CLIPSIBLINGS)
	_SkinDockFixed($iCtrl)

	Local $iIndex = UBound($g_aSkinHeaders)
	ReDim $g_aSkinHeaders[$iIndex + 1][7]
	$g_aSkinHeaders[$iIndex][0] = $iCtrl
	$g_aSkinHeaders[$iIndex][1] = 0
	$g_aSkinHeaders[$iIndex][2] = $iW
	$g_aSkinHeaders[$iIndex][3] = $iH
	$g_aSkinHeaders[$iIndex][4] = $sLogoPath
	$g_aSkinHeaders[$iIndex][5] = $sTitle
	$g_aSkinHeaders[$iIndex][6] = $sSubtitle

	__SkinHeaderRender($iIndex)
	Return $iCtrl
EndFunc   ;==>_SkinHeaderCreate


; Новый текст (язык, версия video-compare): ширина контрола пересчитывается
Func _SkinHeaderSetText($iCtrl, $sTitle, $sSubtitle)
	Local $i = __SkinHeaderIndexOf($iCtrl)
	If $i < 0 Then Return
	$g_aSkinHeaders[$i][5] = $sTitle
	$g_aSkinHeaders[$i][6] = $sSubtitle

	Local $iW = __SkinHeaderContentW($sTitle, $sSubtitle)
	If $iW <> $g_aSkinHeaders[$i][2] Then
		$g_aSkinHeaders[$i][2] = $iW
		GUICtrlSetPos($iCtrl, Default, Default, $iW, $g_aSkinHeaders[$i][3])
	EndIf
	__SkinHeaderRender($i)
EndFunc   ;==>_SkinHeaderSetText


Func _SkinHeaderRenderAll()
	For $i = 0 To UBound($g_aSkinHeaders) - 1
		__SkinHeaderRender($i)
	Next
EndFunc   ;==>_SkinHeaderRenderAll


Func _SkinHeaderShutdown()
	For $i = 0 To UBound($g_aSkinHeaders) - 1
		If $g_aSkinHeaders[$i][1] Then _WinAPI_DeleteObject($g_aSkinHeaders[$i][1])
		$g_aSkinHeaders[$i][1] = 0
	Next
EndFunc   ;==>_SkinHeaderShutdown


; ============================================================
; Внутреннее
; ============================================================

Func __SkinHeaderIndexOf($iCtrl)
	For $i = 0 To UBound($g_aSkinHeaders) - 1
		If $g_aSkinHeaders[$i][0] = $iCtrl Then Return $i
	Next
	Return -1
EndFunc   ;==>__SkinHeaderIndexOf


; Ширина по контенту: логотип и самая широкая из двух строк
Func __SkinHeaderContentW($sTitle, $sSubtitle)
	Local $iTextX = $gc_iSkinHeaderPadL + $gc_iSkinHeaderLogo + $gc_iSkinHeaderTextGap
	Local $iTitleW = _SkinTextW($sTitle, _SkinFont($g_nSkinSizeTitle, True))
	Local $iSubW = _SkinTextW($sSubtitle, _SkinFont($g_nSkinSizeSmall))
	Local $iW = $iTextX + _Max($iTitleW, $iSubW) + $gc_iSkinHeaderPadR
	Return ($iW > $gc_iSkinHeaderMaxW) ? $gc_iSkinHeaderMaxW : $iW
EndFunc   ;==>__SkinHeaderContentW


Func __SkinHeaderRender($iIndex)
	Local $iW = $g_aSkinHeaders[$iIndex][2], $iH = $g_aSkinHeaders[$iIndex][3]
	If $iW < 1 Or $iH < 1 Then Return

	Local $hGfx
	Local $hCanvas = _SkinCanvas($iW, $iH, _SkinArgb($g_iSkinBg), $hGfx)

	Local $iLogoY = Int(($iH - $gc_iSkinHeaderLogo) / 2)
	_SkinDrawImage($hGfx, $g_aSkinHeaders[$iIndex][4], $gc_iSkinHeaderPadL, $iLogoY, _
			$gc_iSkinHeaderLogo, $gc_iSkinHeaderLogo)

	; Название и подпись – один блок по центру высоты логотипа
	Local $iTextX = $gc_iSkinHeaderPadL + $gc_iSkinHeaderLogo + $gc_iSkinHeaderTextGap
	Local $iTextW = $iW - $iTextX - $gc_iSkinHeaderPadR
	_SkinText($hGfx, $g_aSkinHeaders[$iIndex][5], $iTextX, $iLogoY - 3, $iTextW, 20, _
			_SkinFont($g_nSkinSizeTitle, True), _SkinArgb($g_iSkinText1), 0, 1, 3)
	_SkinText($hGfx, $g_aSkinHeaders[$iIndex][6], $iTextX + 1, $iLogoY + 18, $iTextW, 16, _
			_SkinFont($g_nSkinSizeSmall), _SkinArgb($g_iSkinText3), 0, 1, 3)

	_SkinCanvasApply($g_aSkinHeaders[$iIndex][0], $g_aSkinHeaders[$iIndex][1], $hCanvas, $hGfx)
EndFunc   ;==>__SkinHeaderRender
