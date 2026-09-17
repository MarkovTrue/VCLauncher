#include-once
#include <GDIPlus.au3>
#include <GUIConstantsEx.au3>
#include <WinAPI.au3>
#include <WinAPISysWin.au3>
#include <WindowsConstants.au3>
#include <StaticConstants.au3>
#include <SendMessage.au3>
#include <Math.au3>

; ============================================================
; SkinCore.au3 – общее ядро скинизированных контролов
; ============================================================
; Штатные Edit и Button прямоугольные и под тему не красятся, поэтому вид
; рисуется в GDI+ и кладётся картинкой в Pic. Здесь общее для всех модулей:
; палитра, шрифты, примитивы, кэш иконок, опрос наведения и курсор-рука.
; Модули контролов работают только через этот файл и не знают друг о друге.

; --- Палитра. Заполняет _SkinSetTheme, читают и контролы, и само приложение ---
Global $g_iSkinBg          ; фон окна
Global $g_iSkinCard        ; поверхность карточки/панели
Global $g_iSkinCardBorder
Global $g_iSkinCtrlBg      ; заливка поля ввода и кнопки
Global $g_iSkinCtrlBorder
Global $g_iSkinCtrlUnder   ; усиленная нижняя грань поля (приём Fluent)
Global $g_iSkinText1       ; основной текст
Global $g_iSkinText2       ; подписи строк
Global $g_iSkinText3       ; пояснения под полями
Global $g_iSkinTrack       ; дорожка сегментов
Global $g_iSkinTrackBorder
Global $g_iSkinAccent, $g_iSkinAccentEdge, $g_iSkinAccentHot, $g_iSkinOnAccent
Global $g_iSkinDivider
Global $g_iSkinKeyBg, $g_iSkinKeyBorder
Global $g_iSkinSegSelBg, $g_iSkinSegSelBorder, $g_iSkinSegSelText
Global $g_iSkinHover       ; подсветка контрола под курсором
Global $g_bSkinDark = False

; --- Состояние модуля ---
Global $g_hSkinGui = 0
Global $g_sSkinFont = "Segoe UI", $g_nSkinFontSize = 9
; Типографическая шкала. Считается от кегля приложения в _SkinInit:
; пояснения на ступень мельче основного текста, название шапки – крупный кегль.
Global $g_nSkinSizeSmall = 8, $g_nSkinSizeBody = 9, $g_nSkinSizeTitle = 12.5
Global $g_sSkinIconDir = ""
Global $g_bSkinStarted = False
Global $g_hSkinMeasureBmp = 0, $g_hSkinMeasureGfx = 0
Global $g_hSkinCursorCB = 0
Global $g_mSkinFonts[], $g_mSkinIcons[]
Global $g_aSkinHoverFuncs[0]
Global $g_iSkinHoverCtrl = 0 ; контрол под курсором на текущем такте опроса
Global $g_iSkinGdiLineH = 0  ; кэш _SkinGdiLineHeight: шрифт окна не меняется

; Радиусы скругления: поле и кнопка, дорожка сегментов, карточка.
; Мелкие: на контроле высотой 28 крупное скругление съедает угол.
Global Const $gc_nSkinRadCtrl = 3, $gc_nSkinRadTrack = 4, $gc_nSkinRadCard = 4


; Поднимает GDI+, служебный контекст для измерения текста и опрос наведения.
Func _SkinInit($hGui, $sFont, $nFontSize, $sIconDir)
	$g_hSkinGui = $hGui
	$g_sSkinFont = $sFont
	$g_nSkinFontSize = $nFontSize
	$g_sSkinIconDir = $sIconDir
	$g_nSkinSizeBody = $nFontSize
	$g_nSkinSizeSmall = $nFontSize - 1
	$g_nSkinSizeTitle = $nFontSize + 3.5

	If Not $g_bSkinStarted Then
		_GDIPlus_Startup()
		$g_bSkinStarted = True
	EndIf

	; Ширина текста меряется на служебном контексте: контрола под рукой может не быть
	$g_hSkinMeasureBmp = _GDIPlus_BitmapCreateFromScan0(8, 8)
	$g_hSkinMeasureGfx = _GDIPlus_ImageGetGraphicsContext($g_hSkinMeasureBmp)
	_GDIPlus_GraphicsSetTextRenderingHint($g_hSkinMeasureGfx, 5)

	AdlibRegister("__SkinHoverTick", 70)
EndFunc   ;==>_SkinInit


Func _SkinShutdown()
	AdlibUnRegister("__SkinHoverTick")
	__SkinIconCacheClear()
	__SkinFontCacheClear()

	If $g_hSkinCursorCB Then
		DllCallbackFree($g_hSkinCursorCB)
		$g_hSkinCursorCB = 0
	EndIf
	If $g_hSkinMeasureGfx Then _GDIPlus_GraphicsDispose($g_hSkinMeasureGfx)
	If $g_hSkinMeasureBmp Then _GDIPlus_BitmapDispose($g_hSkinMeasureBmp)
	$g_hSkinMeasureGfx = 0
	$g_hSkinMeasureBmp = 0

	If $g_bSkinStarted Then
		_GDIPlus_Shutdown()
		$g_bSkinStarted = False
	EndIf
EndFunc   ;==>_SkinShutdown


Func _SkinGui()
	Return $g_hSkinGui
EndFunc   ;==>_SkinGui


; Палитра Windows 11 Fluent. $iAccent – акцент 0xRRGGBB, 0 – фирменный синий Fluent.
Func _SkinSetTheme($bDark, $iAccent = 0)
	$g_bSkinDark = $bDark
	If $iAccent = 0 Then $iAccent = 0x0F6CBD

	If $bDark Then
		$g_iSkinBg = 0x202020
		$g_iSkinCard = 0x2B2B2B
		$g_iSkinCardBorder = 0x3A3A3A
		$g_iSkinCtrlBg = 0x333333
		$g_iSkinCtrlBorder = 0x454545
		$g_iSkinCtrlUnder = 0x5A5A5A
		$g_iSkinText1 = 0xFFFFFF
		$g_iSkinText2 = 0xCFCFCF
		$g_iSkinText3 = 0x9B9B9B
		$g_iSkinTrack = 0x2F2F2F
		$g_iSkinTrackBorder = 0x3F3F3F
		$g_iSkinDivider = 0x3D3D3D
		$g_iSkinKeyBg = 0x3A3A3A
		$g_iSkinKeyBorder = 0x4C4C4C
		$g_iSkinSegSelBg = 0x4A4A4A
		$g_iSkinSegSelBorder = 0x5A5A5A
		$g_iSkinSegSelText = 0xFFFFFF
		$g_iSkinHover = 0x3C3C3C
	Else
		$g_iSkinBg = 0xF3F3F3
		$g_iSkinCard = 0xFFFFFF
		$g_iSkinCardBorder = 0xE3E3E3
		$g_iSkinCtrlBg = 0xFFFFFF
		$g_iSkinCtrlBorder = 0xD8D8D8
		$g_iSkinCtrlUnder = 0x9A9A9A
		$g_iSkinText1 = 0x1B1B1B
		$g_iSkinText2 = 0x4A4A4A
		$g_iSkinText3 = 0x7C7C7C
		$g_iSkinTrack = 0xEBEBEB
		$g_iSkinTrackBorder = 0xDADADA
		$g_iSkinDivider = 0xE6E6E6
		$g_iSkinKeyBg = 0xFFFFFF
		$g_iSkinKeyBorder = 0xCECECE
		$g_iSkinSegSelBg = 0xFFFFFF
		$g_iSkinSegSelBorder = 0xC4C4C4
		$g_iSkinSegSelText = 0x1B1B1B
		$g_iSkinHover = 0xE4E4E4
	EndIf

	$g_iSkinAccent = $iAccent
	$g_iSkinAccentEdge = __SkinShade($iAccent, $bDark ? 1.18 : 0.82)
	$g_iSkinAccentHot = __SkinShade($iAccent, $bDark ? 1.12 : 1.10)
	$g_iSkinOnAccent = 0xFFFFFF

	; Иконки кэшируются по цвету, шрифты – нет: при смене темы чистим только иконки
	__SkinIconCacheClear()
EndFunc   ;==>_SkinSetTheme


; ============================================================
; Примитивы рисования
; ============================================================

Func _SkinArgb($iRgb, $iAlpha = 255)
	Return BitAND($iRgb, 0xFFFFFF) + $iAlpha * 0x1000000
EndFunc   ;==>_SkinArgb


; Холст под контрол: прозрачный битмап нужного размера, залитый фоном окна.
; Возвращает GDI+ bitmap, контекст отдаёт через $hGfx.
Func _SkinCanvas($iW, $iH, $iBgArgb, ByRef $hGfx)
	Local $hBitmap = _GDIPlus_BitmapCreateFromScan0($iW, $iH)
	$hGfx = _GDIPlus_ImageGetGraphicsContext($hBitmap)
	_GDIPlus_GraphicsSetSmoothingMode($hGfx, 4)     ; AntiAlias – скругления
	_GDIPlus_GraphicsSetTextRenderingHint($hGfx, 5) ; ClearTypeGridFit
	_GDIPlus_GraphicsSetInterpolationMode($hGfx, 7) ; HighQualityBicubic
	; Half обязателен: при None центр крайнего пикселя лежит на границе холста,
	; правый нижний угол выходит квадратным, а рамка размазывается на две строки
	_GDIPlus_GraphicsSetPixelOffsetMode($hGfx, 4)   ; Half
	_GDIPlus_GraphicsClear($hGfx, $iBgArgb)
	Return $hBitmap
EndFunc   ;==>_SkinCanvas


; Ставит нарисованный холст в Pic и освобождает предыдущий HBITMAP.
Func _SkinCanvasApply($iCtrl, ByRef $hOldBitmap, $hCanvas, $hGfx)
	Local $hBitmap = _GDIPlus_BitmapCreateHBITMAPFromBitmap($hCanvas)
	_GDIPlus_GraphicsDispose($hGfx)
	_GDIPlus_BitmapDispose($hCanvas)
	If Not $hBitmap Then Return

	; Static ComCtl32 v6 хранит копию 32-битной картинки: возвращённая прежняя
	; и наша $hOldBitmap – разные объекты, освобождаются обе
	Local $hPrev = _SendMessage(GUICtrlGetHandle($iCtrl), $STM_SETIMAGE, $IMAGE_BITMAP, $hBitmap, _
			0, "wparam", "handle", "handle")
	If $hPrev Then _WinAPI_DeleteObject($hPrev)

	If $hOldBitmap Then _WinAPI_DeleteObject($hOldBitmap)
	$hOldBitmap = $hBitmap
EndFunc   ;==>_SkinCanvasApply


; Путь прямоугольника со скруглёнными углами. Радиус 0 – прямые углы.
Func _SkinPathRR($nX, $nY, $nW, $nH, $nR)
	Local $hPath = _GDIPlus_PathCreate()
	If $nR <= 0 Then
		_GDIPlus_PathAddRectangle($hPath, $nX, $nY, $nW, $nH)
		Return $hPath
	EndIf

	Local $nD = $nR * 2
	_GDIPlus_PathAddArc($hPath, $nX, $nY, $nD, $nD, 180, 90)
	_GDIPlus_PathAddArc($hPath, $nX + $nW - $nD, $nY, $nD, $nD, 270, 90)
	_GDIPlus_PathAddArc($hPath, $nX + $nW - $nD, $nY + $nH - $nD, $nD, $nD, 0, 90)
	_GDIPlus_PathAddArc($hPath, $nX, $nY + $nH - $nD, $nD, $nD, 90, 90)
	_GDIPlus_PathCloseFigure($hPath)
	Return $hPath
EndFunc   ;==>_SkinPathRR


; Заливка скруглённого прямоугольника. Супервыборка не нужна: снимок окна
; показал одинаково сглаженные четыре угла и при прямой заливке.
Func _SkinFill($hGfx, $nX, $nY, $nW, $nH, $nR, $iArgb)
	If $nW < 1 Or $nH < 1 Then Return
	Local $hPath = _SkinPathRR($nX, $nY, $nW, $nH, $nR)
	Local $hBrush = _GDIPlus_BrushCreateSolid($iArgb)
	_GDIPlus_GraphicsFillPath($hGfx, $hPath, $hBrush)
	_GDIPlus_BrushDispose($hBrush)
	_GDIPlus_PathDispose($hPath)
EndFunc   ;==>_SkinFill


; Рамка внутри заданного прямоугольника: путь сдвинут на полпикселя,
; иначе перо шириной 1 ложится половиной за границу и мылит край.
Func _SkinStroke($hGfx, $nX, $nY, $nW, $nH, $nR, $iArgb, $nWidth = 1)
	If $nW < 1 Or $nH < 1 Then Return
	Local $hPath = _SkinPathRR($nX + 0.5, $nY + 0.5, $nW - 1, $nH - 1, $nR)
	Local $hPen = _GDIPlus_PenCreate($iArgb, $nWidth)
	_GDIPlus_GraphicsDrawPath($hGfx, $hPath, $hPen)
	_GDIPlus_PenDispose($hPen)
	_GDIPlus_PathDispose($hPath)
EndFunc   ;==>_SkinStroke


Func _SkinBox($hGfx, $nX, $nY, $nW, $nH, $nR, $iFillArgb, $iStrokeArgb)
	_SkinFill($hGfx, $nX, $nY, $nW, $nH, $nR, $iFillArgb)
	_SkinStroke($hGfx, $nX, $nY, $nW, $nH, $nR, $iStrokeArgb)
EndFunc   ;==>_SkinBox


Func _SkinLine($hGfx, $nX1, $nY1, $nX2, $nY2, $iArgb)
	Local $hPen = _GDIPlus_PenCreate($iArgb, 1)
	_GDIPlus_GraphicsDrawLine($hGfx, $nX1, $nY1 + 0.5, $nX2, $nY2 + 0.5, $hPen)
	_GDIPlus_PenDispose($hPen)
EndFunc   ;==>_SkinLine


; $iAlign: 0 слева, 1 по центру, 2 справа. $iTrim: 0 без обрезки, 3 многоточие.
Func _SkinText($hGfx, $sText, $nX, $nY, $nW, $nH, $hFont, $iArgb, $iAlign = 0, $iLineAlign = 1, $iTrim = 0)
	Local $tLayout = _GDIPlus_RectFCreate($nX, $nY, $nW, $nH)
	Local $hFormat = _GDIPlus_StringFormatCreate(0x1000) ; NoWrap
	_GDIPlus_StringFormatSetAlign($hFormat, $iAlign)
	_GDIPlus_StringFormatSetLineAlign($hFormat, $iLineAlign)
	If $iTrim Then DllCall("gdiplus.dll", "int", "GdipSetStringFormatTrimming", _
			"handle", $hFormat, "int", $iTrim)

	Local $hBrush = _GDIPlus_BrushCreateSolid($iArgb)
	_GDIPlus_GraphicsDrawStringEx($hGfx, $sText, $hFont, $tLayout, $hFormat, $hBrush)
	_GDIPlus_BrushDispose($hBrush)
	_GDIPlus_StringFormatDispose($hFormat)
EndFunc   ;==>_SkinText


; Высота строки шрифта окна средствами GDI: ей меряет строки штатный Edit,
; поэтому многострочному полю нужна она, а не высота из GDI+.
Func _SkinGdiLineHeight()
	If $g_iSkinGdiLineH Then Return $g_iSkinGdiLineH
	If Not $g_hSkinGui Then Return 15
	Local $hDC = _WinAPI_GetDC($g_hSkinGui)
	If Not $hDC Then Return 15

	Local $hOld = 0
	Local $hFont = _SendMessage($g_hSkinGui, $WM_GETFONT, 0, 0, 0, "wparam", "lparam", "handle")
	If $hFont Then $hOld = _WinAPI_SelectObject($hDC, $hFont)
	Local $tSize = _WinAPI_GetTextExtentPoint32($hDC, "Ag")
	If $hOld Then _WinAPI_SelectObject($hDC, $hOld)
	_WinAPI_ReleaseDC($g_hSkinGui, $hDC)

	Local $iH = IsDllStruct($tSize) ? DllStructGetData($tSize, "Y") : 0
	If $iH < 1 Then Return 15
	$g_iSkinGdiLineH = $iH
	Return $iH
EndFunc   ;==>_SkinGdiLineHeight


Func _SkinTextW($sText, $hFont)
	If Not $g_hSkinMeasureGfx Then Return 0
	Local $tLayout = _GDIPlus_RectFCreate(0, 0, 4000, 100)
	Local $hFormat = _GDIPlus_StringFormatCreate(0x1000)
	Local $aInfo = _GDIPlus_GraphicsMeasureString($g_hSkinMeasureGfx, $sText, $hFont, $tLayout, $hFormat)
	_GDIPlus_StringFormatDispose($hFormat)
	If Not IsArray($aInfo) Then Return 0
	Return Ceiling(DllStructGetData($aInfo[0], "Width"))
EndFunc   ;==>_SkinTextW


; Шрифт из кэша. $bSemi – полужирное начертание Segoe UI Semibold, если оно есть.
Func _SkinFont($nSize = 0, $bSemi = False)
	If $nSize = 0 Then $nSize = $g_nSkinFontSize
	Local $sKey = $nSize & "|" & ($bSemi ? 1 : 0)
	; В кэше лежит пара [шрифт, семейство]: семейство нужно живым, пока жив шрифт
	If MapExists($g_mSkinFonts, $sKey) Then
		Local $aCached = $g_mSkinFonts[$sKey]
		Return $aCached[0]
	EndIf

	Local $sName = $bSemi ? "Segoe UI Semibold" : $g_sSkinFont
	Local $hFamily = _GDIPlus_FontFamilyCreate($sName)
	If @error Or Not $hFamily Then $hFamily = _GDIPlus_FontFamilyCreate($g_sSkinFont)
	Local $hFont = _GDIPlus_FontCreate($hFamily, $nSize, 0)

	Local $aPair[2] = [$hFont, $hFamily]
	$g_mSkinFonts[$sKey] = $aPair
	Return $hFont
EndFunc   ;==>_SkinFont


; ============================================================
; Иконки
; ============================================================

; Иконка PNG, перекрашенная в $iRgb, из кэша. Имя – без расширения.
Func _SkinIcon($sName, $iSize, $iRgb)
	Local $sKey = $sName & "|" & $iSize & "|" & Hex($iRgb, 6)
	If MapExists($g_mSkinIcons, $sKey) Then Return $g_mSkinIcons[$sKey]

	Local $sPath = $g_sSkinIconDir & "\" & $sName & ".png"
	If Not FileExists($sPath) Then Return 0
	Local $hSrc = _GDIPlus_ImageLoadFromFile($sPath)
	If @error Or Not $hSrc Then Return 0

	Local $hTinted = __SkinTintCopy($hSrc, $iSize, $iRgb)
	_GDIPlus_ImageDispose($hSrc)
	If Not $hTinted Then Return 0

	$g_mSkinIcons[$sKey] = $hTinted
	Return $hTinted
EndFunc   ;==>_SkinIcon


Func _SkinDrawIcon($hGfx, $sName, $iSize, $iRgb, $nX, $nY)
	Local $hIcon = _SkinIcon($sName, $iSize, $iRgb)
	If Not $hIcon Then Return
	_GDIPlus_GraphicsDrawImage($hGfx, $hIcon, $nX, $nY)
EndFunc   ;==>_SkinDrawIcon


; Рисует PNG как есть, без перекраски (логотип).
Func _SkinDrawImage($hGfx, $sPath, $nX, $nY, $nW, $nH)
	If Not FileExists($sPath) Then Return
	Local $hImg = _GDIPlus_ImageLoadFromFile($sPath)
	If @error Or Not $hImg Then Return
	_GDIPlus_GraphicsDrawImageRect($hGfx, $hImg, $nX, $nY, $nW, $nH)
	_GDIPlus_ImageDispose($hImg)
EndFunc   ;==>_SkinDrawImage


; ============================================================
; Наведение и курсор
; ============================================================

; Модуль отдаёт имя своей функции опроса наведения, ядро зовёт все одним таймером
Func _SkinHoverRegister($sFunc)
	Local $iIndex = UBound($g_aSkinHoverFuncs)
	ReDim $g_aSkinHoverFuncs[$iIndex + 1]
	$g_aSkinHoverFuncs[$iIndex] = $sFunc
EndFunc   ;==>_SkinHoverRegister


; Контрол под курсором на текущем такте опроса, общий для всех модулей
Func _SkinCtrlUnderCursor()
	Return $g_iSkinHoverCtrl
EndFunc   ;==>_SkinCtrlUnderCursor


; X курсора относительно левого края контрола, -1 если курсор вне контрола
Func _SkinCursorLocalX($iCtrl)
	Local $hGui = __SkinCtrlGui($iCtrl)
	Local $aPos = ControlGetPos($hGui, "", $iCtrl)
	If @error Or Not IsArray($aPos) Then Return -1
	Local $aCursor = GUIGetCursorInfo($hGui)
	If @error Or Not IsArray($aCursor) Then Return -1
	Return $aCursor[0] - $aPos[0]
EndFunc   ;==>_SkinCursorLocalX


; Отправляет контрол в самый низ порядка перекрытия. Созданные раньше контролы
; AutoIt кладёт поверх поздних, и подложка (рамка поля, шапка) перехватывала бы
; клики у лежащих на ней Edit и кнопок. Звать после создания всего, что на подложке.
Func _SkinSendToBack($iCtrl)
	_WinAPI_SetWindowPos(GUICtrlGetHandle($iCtrl), $HWND_BOTTOM, 0, 0, 0, 0, _
			BitOR($SWP_NOSIZE, $SWP_NOMOVE, $SWP_NOACTIVATE))
EndFunc   ;==>_SkinSendToBack


; Текущий размер контрола [ширина, высота]: после ресайза его задал докинг
Func _SkinCtrlSize($iCtrl)
	Local $aSize[2] = [0, 0]
	Local $aPos = ControlGetPos(__SkinCtrlGui($iCtrl), "", $iCtrl)
	If @error Or Not IsArray($aPos) Then Return $aSize
	$aSize[0] = $aPos[2]
	$aSize[1] = $aPos[3]
	Return $aSize
EndFunc   ;==>_SkinCtrlSize


; Докинг «стоит на месте». Назначивший свой докинг обязан после ресайза
; вызвать _Skin*SyncSize, иначе картинка останется прежнего размера.
Func _SkinDockFixed($iCtrl)
	GUICtrlSetResizing($iCtrl, $GUI_DOCKLEFT + $GUI_DOCKTOP + $GUI_DOCKWIDTH + $GUI_DOCKHEIGHT)
EndFunc   ;==>_SkinDockFixed


; Курсор-рука над контролом. GUICtrlSetCursor на Pic не действует,
; поэтому на WM_SETCURSOR отвечает subclass.
Func _SkinHandCursor($iCtrl)
	If $g_hSkinCursorCB = 0 Then
		$g_hSkinCursorCB = DllCallbackRegister("__SkinCursorProc", "lresult", _
				"hwnd;uint;wparam;lparam;uint_ptr;dword_ptr")
	EndIf
	_WinAPI_SetWindowSubclass(GUICtrlGetHandle($iCtrl), DllCallbackGetPtr($g_hSkinCursorCB), 1, 0)
EndFunc   ;==>_SkinHandCursor


Func _SkinHandCursorRemove($iCtrl)
	If Not $g_hSkinCursorCB Then Return
	_WinAPI_RemoveWindowSubclass(GUICtrlGetHandle($iCtrl), DllCallbackGetPtr($g_hSkinCursorCB), 1)
EndFunc   ;==>_SkinHandCursorRemove


; ============================================================
; Внутреннее
; ============================================================

; Окно, которому принадлежит контрол. $g_hSkinGui – только главное,
; а скин живёт и в окне настроек.
Func __SkinCtrlGui($iCtrl)
	Return _WinAPI_GetParent(GUICtrlGetHandle($iCtrl))
EndFunc   ;==>__SkinCtrlGui


; Такт опроса наведения: контрол под курсором ищется один раз на все модули
Func __SkinHoverTick()
	$g_iSkinHoverCtrl = __SkinFindCtrlUnderCursor()
	For $i = 0 To UBound($g_aSkinHoverFuncs) - 1
		Call($g_aSkinHoverFuncs[$i])
	Next
EndFunc   ;==>__SkinHoverTick


; Опрашивается окно, реально лежащее под курсором: опрос одного главного не видел бы
; контролы настроек и подсвечивал бы контрол под перекрывающим его окном.
Func __SkinFindCtrlUnderCursor()
	; GetCursorPos напрямую: _WinAPI_GetMousePos на каждом такте дёргал бы Opt дважды
	Local $tPoint = DllStructCreate($tagPOINT)
	DllCall("user32.dll", "bool", "GetCursorPos", "struct*", $tPoint)
	If @error Then Return 0
	Local $hWnd = _WinAPI_WindowFromPoint($tPoint)
	If Not $hWnd Then Return 0
	Local $aCursor = GUIGetCursorInfo(_WinAPI_GetAncestor($hWnd, $GA_ROOT))
	If @error Or Not IsArray($aCursor) Then Return 0
	Return $aCursor[4]
EndFunc   ;==>__SkinFindCtrlUnderCursor


Func __SkinCursorProc($hWnd, $iMsg, $wParam, $lParam, $iId, $dwData)
	#forceref $iId, $dwData
	If $iMsg = $WM_SETCURSOR Then
		_WinAPI_SetCursor(_WinAPI_LoadCursor(0, $IDC_HAND))
		Return 1
	EndIf
	Return _WinAPI_DefSubclassProc($hWnd, $iMsg, $wParam, $lParam)
EndFunc   ;==>__SkinCursorProc


; Копия картинки в один цвет: RGB зануляется матрицей и задаётся строкой
; смещения, альфа исходника сохраняется, и силуэт остаётся сглаженным.
Func __SkinTintCopy($hImage, $iSize, $iColor)
	Local $hBmp = _GDIPlus_BitmapCreateFromScan0($iSize, $iSize)
	If @error Or Not $hBmp Then Return 0
	Local $hGfx = _GDIPlus_ImageGetGraphicsContext($hBmp)
	; PNG уже нарисован в целевом размере: копируем пиксель в пиксель
	_GDIPlus_GraphicsSetInterpolationMode($hGfx, 5) ; NearestNeighbor
	_GDIPlus_GraphicsSetPixelOffsetMode($hGfx, 3) ; None

	Local $nR = BitAND(BitShift($iColor, 16), 0xFF) / 255
	Local $nG = BitAND(BitShift($iColor, 8), 0xFF) / 255
	Local $nB = BitAND($iColor, 0xFF) / 255

	Local $tMatrix = DllStructCreate("float m[25]")
	DllStructSetData($tMatrix, "m", 0, 1)  ; R' = 0
	DllStructSetData($tMatrix, "m", 0, 7)  ; G' = 0
	DllStructSetData($tMatrix, "m", 0, 13) ; B' = 0
	DllStructSetData($tMatrix, "m", 1, 19) ; A' = A
	DllStructSetData($tMatrix, "m", $nR, 21)
	DllStructSetData($tMatrix, "m", $nG, 22)
	DllStructSetData($tMatrix, "m", $nB, 23)
	DllStructSetData($tMatrix, "m", 1, 25)

	Local $aAttr = DllCall("gdiplus.dll", "int", "GdipCreateImageAttributes", "handle*", 0)
	If @error Or $aAttr[0] <> 0 Then
		_GDIPlus_GraphicsDispose($hGfx)
		_GDIPlus_BitmapDispose($hBmp)
		Return 0
	EndIf
	Local $hAttr = $aAttr[1]
	DllCall("gdiplus.dll", "int", "GdipSetImageAttributesColorMatrix", "handle", $hAttr, _
			"int", 0, "bool", True, "struct*", $tMatrix, "ptr", 0, "int", 0)
	DllCall("gdiplus.dll", "int", "GdipDrawImageRectRectI", "handle", $hGfx, "handle", $hImage, _
			"int", 0, "int", 0, "int", $iSize, "int", $iSize, _
			"int", 0, "int", 0, "int", _GDIPlus_ImageGetWidth($hImage), "int", _GDIPlus_ImageGetHeight($hImage), _
			"int", 2, "handle", $hAttr, "ptr", 0, "ptr", 0)
	DllCall("gdiplus.dll", "int", "GdipDisposeImageAttributes", "handle", $hAttr)
	_GDIPlus_GraphicsDispose($hGfx)
	Return $hBmp
EndFunc   ;==>__SkinTintCopy


; Осветление ($nFactor > 1) или затемнение ($nFactor < 1) цвета
Func __SkinShade($iRgb, $nFactor)
	Local $iR = _Min(255, Int(BitAND(BitShift($iRgb, 16), 0xFF) * $nFactor))
	Local $iG = _Min(255, Int(BitAND(BitShift($iRgb, 8), 0xFF) * $nFactor))
	Local $iB = _Min(255, Int(BitAND($iRgb, 0xFF) * $nFactor))
	Return BitShift($iR, -16) + BitShift($iG, -8) + $iB
EndFunc   ;==>__SkinShade


Func __SkinIconCacheClear()
	For $sKey In MapKeys($g_mSkinIcons)
		If $g_mSkinIcons[$sKey] Then _GDIPlus_BitmapDispose($g_mSkinIcons[$sKey])
	Next
	Local $mEmpty[]
	$g_mSkinIcons = $mEmpty
EndFunc   ;==>__SkinIconCacheClear


Func __SkinFontCacheClear()
	For $sKey In MapKeys($g_mSkinFonts)
		Local $aPair = $g_mSkinFonts[$sKey]
		If $aPair[0] Then _GDIPlus_FontDispose($aPair[0])
		If $aPair[1] Then _GDIPlus_FontFamilyDispose($aPair[1])
	Next
	Local $mEmpty[]
	$g_mSkinFonts = $mEmpty
EndFunc   ;==>__SkinFontCacheClear
