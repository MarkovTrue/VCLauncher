#include-once

; ============================================================
; FontHelper.au3 — утилиты для проверки и выбора шрифтов
; ============================================================


; Проверяет, установлен ли шрифт с именем $sName в системе (реестр).
; Возвращает True / False.
Func _FontExists($sName)
	Local Const $sKey = "HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Fonts"
	Local $i = 1, $sValName
	While True
		$sValName = RegEnumVal($sKey, $i)
		If @error Then ExitLoop
		; Значения вида "Segoe UI (TrueType)", "Segoe UI Bold (TrueType)" и т.п.
		If StringRegExp($sValName, "(?i)^\Q" & $sName & "\E(\s|\(|$)") Then Return True
		$i += 1
	WEnd
	Return False
EndFunc   ;==>_FontExists


; Находит первый установленный шрифт из массива $aCandidates.
; Если ни один не найден — возвращает $sDefault.
Func _FontFindBest(Const ByRef $aCandidates, $sDefault = "MS Shell Dlg 2")
	For $sCand In $aCandidates
		If _FontExists($sCand) Then Return $sCand
	Next
	Return $sDefault
EndFunc   ;==>_FontFindBest


; Выбирает лучший шрифт из кандидатов и применяет к окну $hGui.
; Возвращает имя выбранного шрифта.
Func _FontApply($hGui, $iSize = 9, $iWeight = 400, $iAttrib = 0)
	Local $aCandidates[3] = ["Segoe UI", "Tahoma", "MS Shell Dlg 2"]
	Local $sFont = _FontFindBest($aCandidates)
	GUISetFont($iSize, $iWeight, $iAttrib, $sFont, $hGui)
	Return $sFont
EndFunc   ;==>_FontApply
