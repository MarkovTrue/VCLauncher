#pragma compile(Out, #Build\VCLauncher.exe)
#pragma compile(Icon, Assets\Icons\Icon.ico)
#pragma compile(ProductName, VCLauncher)
#pragma compile(FileDescription, Side by side video comparison tool)
#pragma compile(FileVersion, 1.11.0.0)
#pragma compile(x64, true)

#NoTrayIcon

#include <GUIConstantsEx.au3>
#include <WindowsConstants.au3>
#include <FileConstants.au3>
#include <WinAPI.au3>
#include <Math.au3>
#include <ColorConstants.au3>
#include <EditConstants.au3>
#include <ButtonConstants.au3>
#include <ProcessConstants.au3>
#include <SecurityConstants.au3>
#include <ImageListConstants.au3>
#include <AutoItConstants.au3>
#include <StaticConstants.au3>
#include <StringConstants.au3>
#include <InetConstants.au3>
#include <Date.au3>

#include "Include\AppConstants.au3"
#include "Include\GUIDarkTheme.au3"
#include "Include\FontHelper.au3"

; Скинизированные контролы: вид рисуется в GDI+, штатные Edit и Button
; прямоугольные и под тему не красятся. Модуль на компонент.
#include "Controls\SkinCore.au3"
#include "Controls\SkinButton.au3"
#include "Controls\SkinInput.au3"
#include "Controls\SkinSegment.au3"
#include "Controls\SkinHeader.au3"
#include "Controls\SkinHotkeys.au3"

Opt("GUIOnEventMode", 1)

; Без DPI awareness Windows на масштабе не 100% растягивает окно битмапом,
; и углы скруглений после даунскейла DWM выходят разными. Ставится до первого окна.
DllCall("user32.dll", "bool", "SetProcessDPIAware")

; ============================================================
; Константы ($gc_sAppName и сетевые URL – в Include\AppConstants.au3)
; ============================================================
Global Const $gc_sPathIni = @ScriptDir & '\VCLauncher.ini'
Global Const $gc_sPathCache = @ScriptDir & '\VCLauncher.cache'
Global Const $gc_aSupportedExtensions[] = [ _
        "mp4", "m4v", "mov", "mkv", "webm", "avi", "wmv", "asf", "flv", "f4v", "mpg", "mpeg", "mp2", "m2v", _
        "ts", "m2ts", "mts", "mxf", "vob", "3gp", "3g2", "ogv", "ogg", "dv", "divx", "rm", "rmvb", "gif", "vpy"]
; Минимальная ширина: в неё обязаны влезть три колонки шпаргалки
Global Const $gc_iGuiWidth = 560

; --- Метрика тела окна (Fluent) ---
; Поле у краёв 16, зазор 12, строка 28: в ней по центру встают и сегменты, и поле ввода
Global Const $gc_iPadX = 16, $gc_iGap = 12, $gc_iRowH = 28, $gc_iIconBtn = 28
Global Const $gc_iLabelW = 76 ; колонка подписей, по самой длинной («Смещение»)
Global Const $gc_iCtrlX = $gc_iPadX + $gc_iLabelW + $gc_iGap ; левый край всех контролов
; Пояснение под строкой: отступ от верха строки и высота
Global Const $gc_iSubTop = 32, $gc_iSubH = 14
; Шаг строк «Видео 1» → «Видео 2»: строка, пояснение и кнопка обмена между ними
Global Const $gc_iRowGap = 70
; Зазор от поля видео до «…» равен вертикальному между «…» и «⇅», тот задан шагом строк
Global Const $gc_iBtnColGap = Int(($gc_iRowGap - $gc_iIconBtn * 2) / 2)
Global Const $gc_iBlockGap = 14 ; воздух над и под разделителем блоков
Global Const $gc_iCommandH = 60, $gc_iCompareBtnW = 104
Global Const $gc_iHeaderBtn = 32 ; кнопки шапки: шпаргалка и настройки
; Ширина строки подсказки в символах: длиннее – подсказка тянется полосой
Global Const $gc_iTipWidth = 46

Global Const $gc_iUpdateTimeoutMs = 30000 ; зависшая загрузка версии обрывается

; Цвета статуса поиска сдвига – единственное цветное место в окне
Global Const $gc_iClrOkDark = 0x4CC2FF, $gc_iClrOkLight = 0x0A66C2
Global Const $gc_iClrWarnDark = 0xFFCC66, $gc_iClrWarnLight = 0x9A6700
Global Const $gc_iClrErrDark = 0xFF8A8A, $gc_iClrErrLight = 0xC42B1C

; Кадры спиннера статуса поиска
Global Const $gc_aSpinnerFrames[4] = [ChrW(0x25D0), ChrW(0x25D3), ChrW(0x25D1), ChrW(0x25D2)]
; Своё сообщение: перерисовка скина после докинга, см. _OnEvent_WM_SIZE
Global Const $gc_iWmSkinSync = $WM_APP + 1

; Ini и cache создаются до первого чтения настроек
_EnsureIniDefaults()
_EnsureUtf16File($gc_sPathCache)

; ============================================================
; Глобальные переменные
; ============================================================
; Пути к инструментам (уточняет _ResolveToolPaths). ffmpeg берём из комплекта
; video-compare: shared-сборка на тех же av*-DLL, свой статический не нужен.
Global $g_sPathVideoCompare = @ScriptDir & '\video-compare.exe'
Global $g_sPathSync = @ScriptDir & '\Sync\dist\Sync.exe'
Global $g_sPathFFmpeg = @ScriptDir & '\ffmpeg.exe'
Global $g_sVcVersion = "" ; версия video-compare для шапки, см. _InitVcVersion

; Укладка кадров на холст, общая для обоих режимов сравнения:
; native – 1:1, меньший кадр по центру большего (--conversion-fit native);
; crop – масштаб к общему размеру и обрезка лишней высоты фильтром crop.
Global $g_sFit = IniRead($gc_sPathIni, "Settings", "Fit", "native")

; Поиск сдвига (Sync.exe): пропуск от начала (--skip), таймаут процесса,
; порог подозрительного сдвига для перепроверки по видео и допуск совпадения
Global $g_iSyncSkipSec = Int(IniRead($gc_sPathIni, "Settings", "SyncSkipSec", 300))
Global $g_iSyncTimeoutSec = Int(IniRead($gc_sPathIni, "Settings", "SyncTimeoutSec", 60))
Global $g_iSyncSuspectMs = Int(IniRead($gc_sPathIni, "Settings", "SyncSuspectMs", 10000))
Global $g_iSyncVerifyTolMs = Int(IniRead($gc_sPathIni, "Settings", "SyncVerifyTolMs", 500))

; Главное окно
Global $g_hGui, $g_iHeaderPic, $g_iButtonSettings, $g_iButtonKeys, $g_iHotkeysPanel
Global $g_iLabel1, $g_iLabel2, $g_iInput1, $g_iInput2, $g_iLabelInfo1, $g_iLabelInfo2
Global $g_iButtonChoose1, $g_iButtonChoose2, $g_iButtonSwap
Global $g_iLabelOffset, $g_iSegOffset, $g_iInputOffset, $g_iLabelOffsetUnits, $g_iLabelSyncStatus
Global $g_iLabelCompare, $g_iSegCompare, $g_iSegFit, $g_iLabelModeDesc, $g_iLabelFitDesc
Global $g_iLabelCommand, $g_iEditCommand, $g_iButtonCompare, $g_iLabelCommandHint
Global $g_iSeparatorTop, $g_iSepFiles, $g_iSepCompare
Global $g_sAppFont = "MS Shell Dlg 2", $g_iAppFontSize = 9

; Размеры: расчётная высота без добавки пользователя, клиентская область и добавка
; сверх расчётной (она целиком уходит полю команды). Поля рамки меряет
; _FitWindowToContent; до замера нули, иначе WM_GETMINMAXINFO ужмёт новое окно.
Global $g_iGuiHeight = 480, $g_iClientW = $gc_iGuiWidth, $g_iClientH = 480, $g_iExtraH = 0
Global $g_iFrameDX = 0, $g_iFrameDY = 0
Global $g_bHotkeysOpen = (IniRead($gc_sPathIni, "Settings", "Hotkeys", "0") = "1")
Global $g_iHotkeysH = 0
Global $g_bSkinSyncPosted = False

; Окно настроек
Global $g_hSettingsGui = 0
Global $g_iSettingsSegLang = 0, $g_iSettingsSegTheme = 0, $g_iSettingsSegUpdates = 0
Global $g_iSettingsLabelLang = 0, $g_iSettingsLabelTheme = 0, $g_iSettingsLabelUpdates = 0
Global $g_iSettingsLabelUpdateInfo = 0 ; ссылка «Доступна версия X.XX» или дата проверки
Global $g_iSettingsSep = 0
Global $g_iSettingsButtonClearCache = 0, $g_iSettingsButtonCheckNow = 0, $g_iSettingsButtonOk = 0
Global $g_hLinkSubclassCB = 0 ; курсор-рука над ссылкой обновления

; Обновления
Global $g_sAvailableVersion = "" ; найденная версия новее текущей
Global $g_bPulseActive = False, $g_bPulsePhase = False ; мигание шестерёнки
; Фоновая загрузка версии: handle InetGet, временный файл, мигать ли при находке, старт
Global $g_hUpdateDownload = 0, $g_sUpdateTemp = "", $g_bUpdatePulse = False, $g_hUpdateTimer = 0

; Кеш в памяти: разрешения и сдвиги. Ключ включает mtime, заменённый файл его обходит.
Global $g_oCache[]

; Последний поиск сдвига: для статуса и MsgBox с деталями
Global $g_sLastSyncStatus = "" ; "" | NOTRUN | MANUAL | WORKING | OK | NOMATCH | TIMEOUT | ERROR
Global $g_iLastSyncOffset = 0
Global $g_sLastSyncCmd = ""
Global $g_sLastSyncOutput = "" ; stdout + stderr Sync.exe
Global $g_iSpinnerIdx = 0

; Локализация и тема
Global $g_sLangFile = "", $g_sCurrentLang = "", $g_mLang[]
Global $g_sTheme = "Light" ; Light | Dark
Global $g_bDarkMode = False
Global $g_bThemeInitialized = False ; UDF-тема уже применялась
Global $g_bAppliedDark = False ; тема, реально применённая к GUI
Global $g_iClrBg, $g_iClrFg, $g_iClrInfo, $g_iClrSep
; Подсветка поля при drag&drop. WM_CTLCOLOREDIT забирает UDF темы,
; поэтому заливку даёт свой делегат _OnEvent_WM_CTLCOLOREDIT.
Global $g_hDragHiCtrl = 0 ; HWND подсвечиваемого поля
Global $g_hBrushDrag = 0  ; кисть подсветки, создаётся лениво

; Выбранные файлы. Относительный путь в ini считается от папки скрипта
Global $g_sVideoFile1 = _NormalizePath(IniRead($gc_sPathIni, "LastDirs", "Video1", ""))
Global $g_sVideoFile2 = _NormalizePath(IniRead($gc_sPathIni, "LastDirs", "Video2", ""))


_ResolveToolPaths()
_InitLanguage()
_InitTheme()

_CheckToolExists($g_sPathVideoCompare, "video-compare.exe")
_CheckToolExists($g_sPathSync, "Sync.exe")
_CheckToolExists($g_sPathFFmpeg, "ffmpeg.exe")

_InitVcVersion()

OnAutoItExitRegister("_Cleanup")

_MainGUI()
_DefineEvents()

; Проверка обновлений в фоне, окно уже показано
_CheckUpdates()

While 1
	Sleep(50)
WEnd


Func _MainGUI()
	; WS_CLIPCHILDREN: без него заливка фона на ресайзе ложится поверх контролов
	; и окно мерцает. Держится на непрозрачном фоне всех подписей (_ApplyTheme),
	; зато GUICtrlSetColor подпись больше не перерисовывает (см. _SetSyncStatus).
	$g_hGui = GUICreate($gc_sAppName, $gc_iGuiWidth, $g_iGuiHeight, -1, -1, _
			$WS_SIZEBOX + $WS_SYSMENU + $WS_MINIMIZEBOX + $WS_CLIPCHILDREN, $WS_EX_ACCEPTFILES)
	_FitWindowToContent() ; размеры GUICreate внешние, раскладка считается от клиентских
	$g_sAppFont = _FontApply($g_hGui, $g_iAppFontSize)

	; Ядро скина до первого контрола. Палитра здесь первичная, чтобы контролы
	; сразу создавались нужного цвета; окончательную ставит _ApplyTheme.
	_SkinInit($g_hGui, $g_sAppFont, $g_iAppFontSize, @ScriptDir & "\Assets\Icons")
	_SkinSetTheme(_ResolveDarkMode())
	_SkinHoverRegister("_SkinBtnHoverTick")
	_SkinHoverRegister("_SkinInputHoverTick")
	_SkinHoverRegister("_SegHoverTick")

	; --- Шапка: логотип, название, версия video-compare ---
	$g_iHeaderPic = _SkinHeaderCreate(0, 0, _
			@ScriptDir & "\Assets\Icons\HeaderIcon.png", $gc_sAppName, _HeaderSubtitle())

	Local $iBx = $gc_iGuiWidth - $gc_iPadX - $gc_iHeaderBtn
	$g_iButtonSettings = _SkinBtnCreate("", "Settings", 20, $iBx, 14, _
			$gc_iHeaderBtn, $gc_iHeaderBtn, $SKINBTN_SUBTLE)
	_SetTip($g_iButtonSettings, Lang("GUI", "Settings", "Settings"))

	$g_iButtonKeys = _SkinBtnCreate("", "Keyboard", 20, $iBx - $gc_iHeaderBtn - 4, 14, _
			$gc_iHeaderBtn, $gc_iHeaderBtn, $SKINBTN_SUBTLE)
	_SetTip($g_iButtonKeys, Lang("Hotkeys", "1", "Video-compare hotkeys"))
	; Шапка – подложка под кнопками, вниз её можно отправить только после них
	_SkinSendToBack($g_iHeaderPic)

	; --- Шпаргалка горячих клавиш: создаётся сразу, раскрывается кнопкой ---
	Local $aRows = _HotkeyRows()
	$g_iHotkeysPanel = _SkinHotkeysCreate($gc_iPadX, $gc_iSkinHeaderH, _
			$gc_iGuiWidth - $gc_iPadX * 2, Lang("Hotkeys", "1", "Video-compare hotkeys"), $aRows)
	$g_iHotkeysH = _SkinHotkeysHeight(UBound($aRows)) + $gc_iGap
	_SkinBtnSetOn($g_iButtonKeys, $g_bHotkeysOpen)
	If Not $g_bHotkeysOpen Then _SkinHotkeysSetShow($g_iHotkeysPanel, False)

	$g_iSeparatorTop = GUICtrlCreateLabel("", 0, 0, $gc_iGuiWidth, 1)

	; --- Тело окна. Координаты раздаёт _LayoutBody, здесь только создание ---
	$g_iLabel1 = GUICtrlCreateLabel(Lang("GUI", "File1", "File 1"), 0, 0, 10, 16)
	$g_iInput1 = _SkinInputCreate(0, 0, 100, $gc_iRowH)
	_AcceptVideoDrop($g_iInput1)
	_ShowVideoFile($g_iInput1, $g_sVideoFile1)
	$g_iButtonChoose1 = _SkinBtnCreate("", "Ellipsis", 18, 0, 0, $gc_iIconBtn, $gc_iIconBtn)
	_SetTip($g_iButtonChoose1, Lang("GUI", "Choose", "Choose"))
	$g_iLabelInfo1 = GUICtrlCreateLabel(Lang("GUI", "FileNotSelected", "File not selected"), 0, 0, 10, $gc_iSubH)

	; Кнопка обмена встаёт ровно посередине между полями
	$g_iButtonSwap = _SkinBtnCreate("", "Swap", 18, 0, 0, $gc_iIconBtn, $gc_iIconBtn)
	_SetTip($g_iButtonSwap, Lang("GUI", "SwapTip", "Swap files"))

	$g_iLabel2 = GUICtrlCreateLabel(Lang("GUI", "File2", "File 2"), 0, 0, 10, 16)
	$g_iInput2 = _SkinInputCreate(0, 0, 100, $gc_iRowH)
	_AcceptVideoDrop($g_iInput2)
	_ShowVideoFile($g_iInput2, $g_sVideoFile2)
	$g_iButtonChoose2 = _SkinBtnCreate("", "Ellipsis", 18, 0, 0, $gc_iIconBtn, $gc_iIconBtn)
	_SetTip($g_iButtonChoose2, Lang("GUI", "Choose", "Choose"))
	$g_iLabelInfo2 = GUICtrlCreateLabel(Lang("GUI", "FileNotSelected", "File not selected"), 0, 0, 10, $gc_iSubH)

	$g_iSepFiles = GUICtrlCreateLabel("", 0, 0, 10, 1)

	; --- Смещение: режим сегментом, значение в поле ---
	$g_iLabelOffset = GUICtrlCreateLabel(Lang("GUI", "Offset", "Offset"), 0, 0, 10, 16)
	$g_iSegOffset = _SegCreate(Lang("GUI", "OffsetAuto", "Auto") & "|" & Lang("GUI", "OffsetManual", "Manual"), _
			"SegAuto|SegManual", 0, 0, $gc_iRowH, 0, "_OnEvent_SegOffsetMode")
	$g_iInputOffset = _SkinInputCreate(0, 0, 88, $gc_iRowH, $ES_RIGHT)
	GUICtrlSetState($g_iInputOffset, $GUI_DISABLE)
	$g_iLabelOffsetUnits = GUICtrlCreateLabel(Lang("GUI", "OffsetUnits", "ms"), 0, 0, 30, 16)

	; Статус поиска кликабелен: по нему открываются детали неудачи
	$g_iLabelSyncStatus = GUICtrlCreateLabel("", 0, 0, 10, $gc_iSubH, $SS_NOTIFY)

	; --- Режим: слева расположение кадров, справа их укладка на холст ---
	$g_iLabelCompare = GUICtrlCreateLabel(Lang("GUI", "CompareMode", "Mode"), 0, 0, 10, 16)
	$g_iSegCompare = _SegCreate(Lang("GUI", "SegDirect", "Direct") & "|" & Lang("GUI", "SegVertical", "Vertical"), _
			"SegDirect|SegVertical", 0, 0, $gc_iRowH, 0, "_OnEvent_SegCompareMode")
	$g_iSegFit = _SegCreate(Lang("GUI", "SegNative", "Native") & "|" & Lang("GUI", "SegCrop", "Cropped"), _
			"SegNative|SegCrop", 0, 0, $gc_iRowH, ($g_sFit = "crop") ? 1 : 0, "_OnEvent_SegFit")
	$g_iLabelModeDesc = GUICtrlCreateLabel("", 0, 0, 10, $gc_iSubH)
	$g_iLabelFitDesc = GUICtrlCreateLabel("", 0, 0, 10, $gc_iSubH)
	; Пояснения несут подсказку с полным текстом: рука подсказывает, что на них стоит навести
	_SkinHandCursor($g_iLabelModeDesc)
	_SkinHandCursor($g_iLabelFitDesc)

	$g_iSepCompare = GUICtrlCreateLabel("", 0, 0, 10, 1)

	; --- Команда и «Сравнить» в одной строке ---
	$g_iLabelCommand = GUICtrlCreateLabel(Lang("GUI", "TabCommand", "Command"), 0, 0, 10, 16)
	$g_iEditCommand = _SkinInputCreate(0, 0, 200, $gc_iCommandH, -1, True)
	$g_iButtonCompare = _SkinBtnCreate(Lang("GUI", "Compare", "Compare"), "CompareDirect", 32, _
			0, 0, $gc_iCompareBtnW, $gc_iCommandH, $SKINBTN_ACCENT)
	$g_iLabelCommandHint = GUICtrlCreateLabel(Lang("GUI", "CommandHint", "Generated automatically. Editable."), _
			0, 0, 10, $gc_iSubH)

	_SetCtrlResizing()
	_UpdateModeDesc()
	_UpdateModeTips()
	_LayoutBody()
	_ApplyTheme()
	_ValidateVideoFiles() ; пути из ini могли устареть, не показываем несуществующие файлы

	GUISetState(@SW_SHOW)
	; Поле «Видео 1» – первый табстоп, при показе Windows выделяет в нём весь текст
	_ResetInputCaret($g_iInput1)

	; Данные подтягиваем, когда окно уже видно
	_TryFillCachedOffset()
	_UpdateFilesInfo()
EndFunc   ;==>_MainGUI


; Привязка контролов к краям окна. Без тегов AutoIt тянет контролы пропорционально
; окну. Картинку скинизированного контрола докинг не перерисовывает, это делает
; _SkinSyncAll после ресайза.
Func _SetCtrlResizing()
	Local $iFixed = $GUI_DOCKLEFT + $GUI_DOCKTOP + $GUI_DOCKWIDTH + $GUI_DOCKHEIGHT
	Local $iStretchH = $GUI_DOCKLEFT + $GUI_DOCKTOP + $GUI_DOCKRIGHT + $GUI_DOCKHEIGHT
	Local $iStretchHV = $GUI_DOCKLEFT + $GUI_DOCKTOP + $GUI_DOCKRIGHT + $GUI_DOCKBOTTOM
	Local $iStretchH_B = $GUI_DOCKLEFT + $GUI_DOCKBOTTOM + $GUI_DOCKRIGHT + $GUI_DOCKHEIGHT
	Local $iFixedRight = $GUI_DOCKRIGHT + $GUI_DOCKTOP + $GUI_DOCKWIDTH + $GUI_DOCKHEIGHT
	Local $iFixedBR = $GUI_DOCKRIGHT + $GUI_DOCKBOTTOM + $GUI_DOCKWIDTH + $GUI_DOCKHEIGHT

	; Тянутся по горизонтали. Шапка остаётся фиксированной: её картинка под контент
	Local $aStretchH[11] = [$g_iHotkeysPanel, $g_iSeparatorTop, $g_iSepFiles, $g_iSepCompare, _
			$g_iInput1, _SkinInputFrame($g_iInput1), $g_iInput2, _SkinInputFrame($g_iInput2), _
			$g_iLabelInfo1, $g_iLabelInfo2, $g_iLabelSyncStatus]
	For $iCtrl In $aStretchH
		GUICtrlSetResizing($iCtrl, $iStretchH)
	Next

	; Поле команды забирает и ширину, и высоту, подсказка под ним держится низа
	GUICtrlSetResizing($g_iEditCommand, $iStretchHV)
	GUICtrlSetResizing(_SkinInputFrame($g_iEditCommand), $iStretchHV)
	GUICtrlSetResizing($g_iLabelCommandHint, $iStretchH_B)

	; Стоят на месте: колонка подписей, переключатели, поле сдвига
	Local $aFixed[13] = [$g_iLabel1, $g_iLabel2, $g_iLabelOffset, $g_iLabelCompare, $g_iLabelCommand, _
			$g_iSegCompare, $g_iSegFit, $g_iSegOffset, _
			$g_iInputOffset, _SkinInputFrame($g_iInputOffset), $g_iLabelOffsetUnits, _
			$g_iLabelModeDesc, $g_iLabelFitDesc]
	For $iCtrl In $aFixed
		GUICtrlSetResizing($iCtrl, $iFixed)
	Next

	; Держатся правого края. Колонки шпаргалки расставляет сама SkinHotkeys
	Local $aRight[5] = [$g_iButtonSettings, $g_iButtonKeys, _
			$g_iButtonChoose1, $g_iButtonSwap, $g_iButtonChoose2]
	For $iCtrl In $aRight
		GUICtrlSetResizing($iCtrl, $iFixedRight)
	Next

	; «Сравнить» – правый нижний угол, размер свой
	GUICtrlSetResizing($g_iButtonCompare, $iFixedBR)
EndFunc   ;==>_SetCtrlResizing


; Раскладка тела окна целиком, от верха тела: он зависит только от того, раскрыта ли
; шпаргалка. Позиции не накапливаются от текущих, поэтому пересчёт всегда даёт тот же вид.
; $bFit = False – высоту окна уже выставил вызывающий.
Func _LayoutBody($bFit = True)
	Local $iRight = $g_iClientW - $gc_iPadX
	Local $iBtnX = $iRight - $gc_iIconBtn     ; правая колонка кнопок-иконок
	Local $iFieldW = $iBtnX - $gc_iBtnColGap - $gc_iCtrlX ; поля тянутся до этой колонки
	Local $iSepW = $g_iClientW - $gc_iPadX * 2

	; Шапка ресайзом не трогается, двигаются только её кнопки у правого края
	Local $iBx = $iRight - $gc_iHeaderBtn
	_SkinBtnSetPos($g_iButtonSettings, $iBx, 14)
	_SkinBtnSetPos($g_iButtonKeys, $iBx - $gc_iHeaderBtn - 4, 14)
	_SkinHotkeysSetPos($g_iHotkeysPanel, $gc_iPadX, $gc_iSkinHeaderH, $iSepW, $g_iHotkeysH - $gc_iGap)

	Local $iTop = $gc_iSkinHeaderH + ($g_bHotkeysOpen ? $g_iHotkeysH : 0)
	GUICtrlSetPos($g_iSeparatorTop, 0, $iTop, $g_iClientW, 1)

	Local $iY = $iTop + 1 + $gc_iBlockGap

	; --- Видео 1 ---
	GUICtrlSetPos($g_iLabel1, $gc_iPadX, $iY + 6, $gc_iLabelW, 16)
	_SkinInputSetPos($g_iInput1, $gc_iCtrlX, $iY, $iFieldW, $gc_iRowH)
	_SkinBtnSetPos($g_iButtonChoose1, $iBtnX, $iY)
	GUICtrlSetPos($g_iLabelInfo1, $gc_iCtrlX + 1, $iY + $gc_iSubTop, $iFieldW, $gc_iSubH)
	_SkinBtnSetPos($g_iButtonSwap, $iBtnX, $iY + $gc_iIconBtn + $gc_iBtnColGap)

	; --- Видео 2 ---
	$iY += $gc_iRowGap
	GUICtrlSetPos($g_iLabel2, $gc_iPadX, $iY + 6, $gc_iLabelW, 16)
	_SkinInputSetPos($g_iInput2, $gc_iCtrlX, $iY, $iFieldW, $gc_iRowH)
	_SkinBtnSetPos($g_iButtonChoose2, $iBtnX, $iY)
	GUICtrlSetPos($g_iLabelInfo2, $gc_iCtrlX + 1, $iY + $gc_iSubTop, $iFieldW, $gc_iSubH)

	$iY = _LayoutSeparator($g_iSepFiles, $iY, $iSepW)

	; --- Смещение ---
	GUICtrlSetPos($g_iLabelOffset, $gc_iPadX, $iY + 6, $gc_iLabelW, 16)
	_SegSetPos($g_iSegOffset, $gc_iCtrlX, $iY)
	Local $iSegW = _SegWidth($g_iSegOffset)
	Local $iOffX = $gc_iCtrlX + $iSegW + $gc_iGap
	Local $iOffW = Int($iSegW / 2) ; поле шириной с одну кнопку переключателя
	_SkinInputSetPos($g_iInputOffset, $iOffX, $iY, $iOffW, $gc_iRowH)
	GUICtrlSetPos($g_iLabelOffsetUnits, $iOffX + $iOffW + 8, $iY + 6, 30, 16)
	GUICtrlSetPos($g_iLabelSyncStatus, $gc_iCtrlX + 1, $iY + $gc_iSubTop, $iFieldW, $gc_iSubH)

	; «Смещение» и «Режим» – одна группа без полосы, с шагом пары полей видео
	$iY += $gc_iRowGap

	; --- Режим: два переключателя подряд, под каждым своё пояснение ---
	GUICtrlSetPos($g_iLabelCompare, $gc_iPadX, $iY + 6, $gc_iLabelW, 16)
	_SegSetPos($g_iSegCompare, $gc_iCtrlX, $iY)
	Local $iModeW = _SegWidth($g_iSegCompare)
	Local $iFitX = $gc_iCtrlX + $iModeW + $gc_iGap
	_SegSetPos($g_iSegFit, $iFitX, $iY)
	GUICtrlSetPos($g_iLabelModeDesc, $gc_iCtrlX + 1, $iY + $gc_iSubTop, $iModeW + 8, $gc_iSubH)
	GUICtrlSetPos($g_iLabelFitDesc, $iFitX + 1, $iY + $gc_iSubTop, $iRight - $iFitX, $gc_iSubH)

	$iY = _LayoutSeparator($g_iSepCompare, $iY, $iSepW)

	; --- Команда и кнопка запуска в одной строке ---
	; Добавленная пользователем высота уходит полю, кнопка держится его нижнего края
	Local $iCmdW = $iRight - $gc_iCompareBtnW - $gc_iGap - $gc_iCtrlX
	Local $iCmdH = $gc_iCommandH + $g_iExtraH

	GUICtrlSetPos($g_iLabelCommand, $gc_iPadX, $iY + 4, $gc_iLabelW, 16)
	_SkinInputSetPos($g_iEditCommand, $gc_iCtrlX, $iY, $iCmdW, $iCmdH)
	_SkinBtnSetPos($g_iButtonCompare, $iRight - $gc_iCompareBtnW, _
			$iY + $iCmdH - $gc_iCommandH, $gc_iCompareBtnW, $gc_iCommandH)
	GUICtrlSetPos($g_iLabelCommandHint, $gc_iCtrlX + 1, $iY + $iCmdH + 4, $iCmdW, $gc_iSubH)

	; Минимальная высота окна без добавки пользователя – её ждёт WM_GETMINMAXINFO
	$g_iGuiHeight = $iY + $gc_iCommandH + 4 + $gc_iSubH + $gc_iBlockGap + 2
	If $bFit Then _FitGuiHeight()
EndFunc   ;==>_LayoutBody


; Разделитель под строкой $iY: по $gc_iBlockGap от низа пояснения до полосы
; и от полосы до следующей строки. Возвращает Y следующей строки.
Func _LayoutSeparator($iCtrl, $iY, $iW)
	Local $iSepY = $iY + $gc_iSubTop + $gc_iSubH + $gc_iBlockGap
	GUICtrlSetPos($iCtrl, $gc_iPadX, $iSepY, $iW, 1)
	Return $iSepY + $gc_iBlockGap + 1
EndFunc   ;==>_LayoutSeparator


; Подгоняет окно под рассчитанную высоту, оставляя левый верхний угол на месте:
; шпаргалка раскрывается вниз, окно не должно прыгать по экрану.
Func _FitGuiHeight()
	Local $aWin = WinGetPos($g_hGui)
	If Not IsArray($aWin) Then Return
	Local $iH = $g_iGuiHeight + $g_iExtraH + $g_iFrameDY
	If $aWin[3] = $iH Then Return
	WinMove($g_hGui, "", $aWin[0], $aWin[1], $aWin[2], $iH)
EndFunc   ;==>_FitGuiHeight


; Подпись под названием в шапке: чем именно управляет лончер
Func _HeaderSubtitle()
	Return "Video-compare" & ($g_sVcVersion <> "" ? " " & $g_sVcVersion : "")
EndFunc   ;==>_HeaderSubtitle


; Строки шпаргалки из языкового файла: «{CTRL}{PLUS}{MINUS} Сдвиг ±10 кадров»
; разбирается на клавиши и подпись. Порядок ключей – порядок строк по трём колонкам:
; вид и режимы, воспроизведение и просмотр, перемотка и сдвиг.
Func _HotkeyRows()
	Local $aKeys[18] = ["2", "3", "7", "8", "9", "6", _
			"18", "19", "14", "17", "16", "15", _
			"4", "5", "13", "10", "11", "12"]
	Local $aRows[UBound($aKeys)][2]

	For $i = 0 To UBound($aKeys) - 1
		Local $sLine = Lang("Hotkeys", $aKeys[$i], "")
		Local $aParts = StringRegExp($sLine, "^((?:\{[^}]+\})+)\s*(.*)$", 1)
		If @error Then
			$aRows[$i][0] = ""
			$aRows[$i][1] = $sLine
			ContinueLoop
		EndIf

		Local $aTokens = StringRegExp($aParts[0], "\{([^}]+)\}", 3)
		Local $sBadges = ""
		For $j = 0 To UBound($aTokens) - 1
			; "ТОКЕН:подпись": по токену ищется иконка клавиши, подпись – для запасного бейджа
			$sBadges &= ($sBadges = "" ? "" : "|") & $aTokens[$j] & ":" & _HotkeyGlyph($aTokens[$j])
		Next
		$aRows[$i][0] = $sBadges
		$aRows[$i][1] = $aParts[1]
	Next
	Return $aRows
EndFunc   ;==>_HotkeyRows


; Надпись на запасном бейдже клавиши по её имени из языкового файла
Func _HotkeyGlyph($sName)
	Switch $sName
		Case "LEFT"
			Return ChrW(0x2190)
		Case "RIGHT"
			Return ChrW(0x2192)
		Case "UP"
			Return ChrW(0x2191)
		Case "DOWN"
			Return ChrW(0x2193)
		Case "PLUS"
			Return "+"
		Case "MINUS"
			Return ChrW(0x2212)
		Case "CTRL"
			Return "Ctrl"
		Case "ALT"
			Return "Alt"
		Case "SHIFT"
			Return "Shift"
	EndSwitch
	Return $sName
EndFunc   ;==>_HotkeyGlyph


; Доводит клиентскую область окна до заданной, сохраняя его центр.
Func _FitGuiToClient($hWnd, $iClientW, $iClientH)
	Local $aWin = WinGetPos($hWnd)
	Local $aClient = WinGetClientSize($hWnd)
	If Not IsArray($aWin) Or Not IsArray($aClient) Then Return

	Local $iW = $iClientW + ($aWin[2] - $aClient[0])
	Local $iH = $iClientH + ($aWin[3] - $aClient[1])
	If $aWin[2] = $iW And $aWin[3] = $iH Then Return

	WinMove($hWnd, "", $aWin[0] + Int(($aWin[2] - $iW) / 2), $aWin[1] + Int(($aWin[3] - $iH) / 2), $iW, $iH)
EndFunc   ;==>_FitGuiToClient


; Доводит клиентскую область до расчётной и запоминает поля рамки: они зависят
; от темы и масштаба Windows, поэтому меряются на живом окне. Вызывается
; до создания контролов, иначе ресайз сдвинул бы их докингом.
Func _FitWindowToContent()
	Local $aWin = WinGetPos($g_hGui)
	Local $aClient = WinGetClientSize($g_hGui)
	If Not IsArray($aWin) Or Not IsArray($aClient) Then Return

	$g_iFrameDX = $aWin[2] - $aClient[0]
	$g_iFrameDY = $aWin[3] - $aClient[1]

	Local $iW = $gc_iGuiWidth + $g_iFrameDX
	Local $iH = $g_iGuiHeight + $g_iFrameDY
	If $aWin[2] = $iW And $aWin[3] = $iH Then Return

	; По центру рабочей области
	Local $tWork = _WinAPI_GetWorkArea()
	Local $iLeft = DllStructGetData($tWork, "Left"), $iTop = DllStructGetData($tWork, "Top")
	$iLeft += Int((DllStructGetData($tWork, "Right") - $iLeft - $iW) / 2)
	$iTop += Int((DllStructGetData($tWork, "Bottom") - $iTop - $iH) / 2)
	WinMove($g_hGui, "", $iLeft, $iTop, $iW, $iH)
EndFunc   ;==>_FitWindowToContent


Func _DefineEvents()
	GUICtrlSetOnEvent($g_iButtonChoose1, "_OnEvent_ButtonChoose")
	GUICtrlSetOnEvent($g_iButtonChoose2, "_OnEvent_ButtonChoose")
	GUICtrlSetOnEvent($g_iButtonSwap, "_OnEvent_ButtonSwap")
	GUICtrlSetOnEvent($g_iButtonCompare, "_OnEvent_ButtonCompare")
	GUICtrlSetOnEvent($g_iButtonSettings, "_OnEvent_ButtonSettings")
	GUICtrlSetOnEvent($g_iButtonKeys, "_OnEvent_ButtonKeys")
	GUICtrlSetOnEvent($g_iLabelSyncStatus, "_OnEvent_LabelSyncStatusClick")
	; Сегменты вешают свои обработчики сами при создании (_SegCreate)
	GUISetOnEvent($GUI_EVENT_CLOSE, "_OnEvent_GUI_EVENT_CLOSE")
	GUISetOnEvent($GUI_EVENT_DROPPED, "_OnEvent_GUI_EVENT_DROPPED")

	GUIRegisterMsg($WM_GETMINMAXINFO, "_OnEvent_WM_GETMINMAXINFO")
	GUIRegisterMsg($WM_COMMAND, "_OnEvent_WM_COMMAND")
	GUIRegisterMsg($WM_DROPFILES, "_OnEvent_WM_DROPFILES")
	GUIRegisterMsg($WM_SIZE, "_OnEvent_WM_SIZE")
	GUIRegisterMsg($WM_EXITSIZEMOVE, "_OnEvent_WM_EXITSIZEMOVE")
	GUIRegisterMsg($gc_iWmSkinSync, "_OnEvent_WM_SKINSYNC")
EndFunc   ;==>_DefineEvents


; Страховка на отпускание рамки: докинг уже применён, размеры окончательные
Func _OnEvent_WM_EXITSIZEMOVE($hWnd, $iMsg, $wParam, $lParam)
	#forceref $iMsg, $wParam, $lParam
	If $hWnd <> $g_hGui Then Return $GUI_RUNDEFMSG
	_SkinSyncAll()
	Return $GUI_RUNDEFMSG
EndFunc   ;==>_OnEvent_WM_EXITSIZEMOVE


Func _OnEvent_WM_GETMINMAXINFO($hWnd, $iMsg, $wParam, $lParam)
	#forceref $iMsg, $wParam
	If $hWnd <> $g_hGui Then Return $GUI_RUNDEFMSG

	Local $tMMI = DllStructCreate( _
			"int reserved1;int reserved2;" & _
			"int MaxSizeX;int MaxSizeY;" & _
			"int MaxPositionX;int MaxPositionY;" & _
			"int MinTrackSizeX;int MinTrackSizeY;" & _
			"int MaxTrackSizeX;int MaxTrackSizeY", $lParam)
	; Меньше расчётного размера контролы наезжают друг на друга
	DllStructSetData($tMMI, "MinTrackSizeX", $gc_iGuiWidth + $g_iFrameDX)
	DllStructSetData($tMMI, "MinTrackSizeY", $g_iGuiHeight + $g_iFrameDY)

	Return 0
EndFunc   ;==>_OnEvent_WM_GETMINMAXINFO


; Контролы двигает докинг (_SetCtrlResizing). Картинки скина перерисовываются
; отложенно: WM_SIZE приходит до докинга, а posted-сообщение – уже после него.
; Без отсрочки Win+стрелка и WinMove оставляли картинки на шаг позади.
Func _OnEvent_WM_SIZE($hWnd, $iMsg, $wParam, $lParam)
	#forceref $iMsg, $wParam
	If $hWnd <> $g_hGui Then Return $GUI_RUNDEFMSG

	Local $iWidth = BitAND($lParam, 0xFFFF)
	Local $iHeight = BitShift($lParam, 16)
	If $iWidth < 100 Or ($iWidth = $g_iClientW And $iHeight = $g_iClientH) Then Return $GUI_RUNDEFMSG

	$g_iClientW = $iWidth
	$g_iClientH = $iHeight
	$g_iExtraH = _Max(0, $g_iClientH - $g_iGuiHeight)

	; Одной синхронизации хватает на серию WM_SIZE
	If Not $g_bSkinSyncPosted Then
		$g_bSkinSyncPosted = True
		_WinAPI_PostMessage($g_hGui, $gc_iWmSkinSync, 0, 0)
	EndIf
	Return $GUI_RUNDEFMSG
EndFunc   ;==>_OnEvent_WM_SIZE


Func _OnEvent_WM_SKINSYNC($hWnd, $iMsg, $wParam, $lParam)
	#forceref $iMsg, $wParam, $lParam
	If $hWnd <> $g_hGui Then Return $GUI_RUNDEFMSG
	$g_bSkinSyncPosted = False
	_SkinSyncAll()
	Return 0
EndFunc   ;==>_OnEvent_WM_SKINSYNC


; Перерисовка картинок скина под размер от докинга. Шапка от ширины окна не зависит
Func _SkinSyncAll()
	_SkinHotkeysSyncSize()
	_SkinInputSyncSize()
	_SkinBtnSyncSize()
EndFunc   ;==>_SkinSyncAll


; Кнопка-клавиатура раскрывает шпаргалку. Состояние помнится в ini
Func _OnEvent_ButtonKeys()
	$g_bHotkeysOpen = Not $g_bHotkeysOpen
	_SkinBtnSetOn($g_iButtonKeys, $g_bHotkeysOpen)
	_SkinHotkeysSetShow($g_iHotkeysPanel, $g_bHotkeysOpen)

	; Сначала высота окна, потом раскладка: иначе докинг сдвинет нижний блок второй раз
	$g_iGuiHeight += $g_bHotkeysOpen ? $g_iHotkeysH : -$g_iHotkeysH
	_FitGuiHeight()
	_LayoutBody(False)
	; Прижатые к низу контролы докинг уже поставил на место, раскладка их не двигает,
	; и без полной перерисовки при скрытии от них оставались обрывки
	_RedrawGui()
	IniWrite($gc_sPathIni, "Settings", "Hotkeys", $g_bHotkeysOpen ? "1" : "0")
EndFunc   ;==>_OnEvent_ButtonKeys


Func _OnEvent_ButtonChoose()
	Local $iButtonID = @GUI_CtrlId
	Local $sTitle, $sIniKey, $iInputCtrl, $sCurrentFile

	If $iButtonID = $g_iButtonChoose1 Then
		$sTitle = Lang("Dialogs", "SelectVideo1", "Select video 1")
		$sIniKey = "Video1"
		$iInputCtrl = $g_iInput1
		$sCurrentFile = $g_sVideoFile1
	Else
		$sTitle = Lang("Dialogs", "SelectVideo2", "Select video 2")
		$sIniKey = "Video2"
		$iInputCtrl = $g_iInput2
		$sCurrentFile = $g_sVideoFile2
	EndIf

	; Слот мог быть сброшен из-за пропавшего файла, стартовую папку тогда берём из ini
	Local $sStartFile = ($sCurrentFile <> "") ? $sCurrentFile _
			: _NormalizePath(IniRead($gc_sPathIni, "LastDirs", $sIniKey, ""))
	Local $sFile = FileOpenDialog($sTitle, _PathGetDir($sStartFile), _GetVideoExtensionsFilter(), _
			$FD_FILEMUSTEXIST, _GetFileName($sCurrentFile))
	If @error Or Not FileExists($sFile) Then Return
	_SetVideoFile($iInputCtrl, $sFile)
EndFunc   ;==>_OnEvent_ButtonChoose


Func _OnEvent_ButtonCompare()
	If _ValidateVideoFiles() Then
		_UpdateFilesInfo()
		Return
	EndIf
	If Not FileExists($g_sVideoFile1) Or Not FileExists($g_sVideoFile2) Then Return

	; На время поиска сдвига и работы video-compare кнопка заблокирована
	_SkinBtnSetEnabled($g_iButtonCompare, False)
	_SkinBtnSetText($g_iButtonCompare, Lang("GUI", "Running", "Running"))

	If _IsAutoOffset() Then
		_SetSyncStatus("WORKING")
		Local $aSync = _GetSyncOffset($g_sVideoFile1, $g_sVideoFile2)
		GUICtrlSetData($g_iInputOffset, ($aSync[0] = "OK") ? $aSync[1] : "")
		_SetSyncStatus($aSync[0], $aSync[1])
		_UpdateCommandField()
	EndIf

	Local $sCmdLine = StringRegExpReplace(GUICtrlRead($g_iEditCommand), "[\r\n]+", " ")
	$sCmdLine = StringStripWS($sCmdLine, $STR_STRIPLEADING + $STR_STRIPTRAILING)
	If $sCmdLine <> "" Then _RunVideoCompare($sCmdLine)

	; Доступность кнопке возвращает _UpdateFilesInfo
	_SkinBtnSetText($g_iButtonCompare, Lang("GUI", "Compare", "Compare"))
	_UpdateFilesInfo()
EndFunc   ;==>_OnEvent_ButtonCompare


Func _OnEvent_ButtonSettings()
	If $g_hSettingsGui <> 0 Then Return
	_StopSettingsPulse() ; пользователь заметил обновление
	_SettingsWindow()
EndFunc   ;==>_OnEvent_ButtonSettings


Func _OnEvent_ButtonSwap()
	_ValidateVideoFiles() ; пропавший файл не переносим в соседний слот

	Local $sTmp = $g_sVideoFile1
	$g_sVideoFile1 = $g_sVideoFile2
	$g_sVideoFile2 = $sTmp

	IniWrite($gc_sPathIni, "LastDirs", "Video1", $g_sVideoFile1)
	IniWrite($gc_sPathIni, "LastDirs", "Video2", $g_sVideoFile2)
	_ShowVideoFile($g_iInput1, $g_sVideoFile1)
	_ShowVideoFile($g_iInput2, $g_sVideoFile2)

	; Сдвиг и его статус инвертируем вместо повторного поиска
	Local $sOffset = GUICtrlRead($g_iInputOffset)
	If $sOffset <> "" Then GUICtrlSetData($g_iInputOffset, -Int($sOffset))
	If $g_sLastSyncStatus = "OK" Then _SetSyncStatus("OK", -$g_iLastSyncOffset)
	_UpdateFilesInfo()
EndFunc   ;==>_OnEvent_ButtonSwap


; Укладка кадров от режима сравнения не зависит и при смене остаётся прежней
Func _OnEvent_SegCompareMode()
	_UpdateModeDesc()
	_UpdateCompareButtonIcon()
	_UpdateCommandField()
EndFunc   ;==>_OnEvent_SegCompareMode


Func _OnEvent_SegFit()
	$g_sFit = (_SegGetSel($g_iSegFit) = 1) ? "crop" : "native"
	IniWrite($gc_sPathIni, "Settings", "Fit", $g_sFit)
	_UpdateModeDesc()
	_UpdateCommandField()
EndFunc   ;==>_OnEvent_SegFit


Func _OnEvent_SegOffsetMode()
	If _IsAutoOffset() Then
		GUICtrlSetState($g_iInputOffset, $GUI_DISABLE)
		_TryFillCachedOffset()
	Else
		GUICtrlSetState($g_iInputOffset, $GUI_ENABLE)
		_SetSyncStatus("MANUAL")
	EndIf
	_UpdateCommandField()
EndFunc   ;==>_OnEvent_SegOffsetMode


Func _OnEvent_GUI_EVENT_CLOSE()
	Exit
EndFunc   ;==>_OnEvent_GUI_EVENT_CLOSE


; Подсветка поля, над которым отпустили файл. Цвет даёт _OnEvent_WM_CTLCOLOREDIT,
; сам файл принимает _OnEvent_GUI_EVENT_DROPPED. Поле ищется по точке броска,
; как и у AutoIt: курсор к этому моменту мог уйти.
Func _OnEvent_WM_DROPFILES($hWnd, $iMsg, $wParam, $lParam)
	#forceref $iMsg, $lParam
	If $hWnd <> $g_hGui Then Return $GUI_RUNDEFMSG

	Local $tPoint = _WinAPI_DragQueryPoint($wParam)
	If @error Then Return $GUI_RUNDEFMSG
	Local $hChild = _WinAPI_ChildWindowFromPointEx($g_hGui, $tPoint, $CWP_SKIPINVISIBLE)
	Local $iInput = $hChild ? _VideoInputOf(_WinAPI_GetDlgCtrlID($hChild)) : 0

	$g_hDragHiCtrl = 0
	If $iInput Then
		$g_hDragHiCtrl = GUICtrlGetHandle($iInput)
		_WinAPI_InvalidateRect($g_hDragHiCtrl)
	EndIf
	AdlibRegister("_RestoreControlsStyle", 100)
	Return $GUI_RUNDEFMSG
EndFunc   ;==>_OnEvent_WM_DROPFILES


; Делегат покраски Edit: UDF темы перехватывает WM_CTLCOLOREDIT глобально,
; поэтому drag-подсветку вставляем здесь, остальные поля отдаём UDF.
Func _OnEvent_WM_CTLCOLOREDIT($hWnd, $iMsg, $wParam, $lParam)
	If $g_hDragHiCtrl <> 0 And $lParam = $g_hDragHiCtrl Then
		If Not $g_hBrushDrag Then $g_hBrushDrag = _WinAPI_CreateSolidBrush(_WinAPI_SwitchColor($COLOR_SKYBLUE))
		_WinAPI_SetBkColor($wParam, _WinAPI_SwitchColor($COLOR_SKYBLUE))
		_WinAPI_SetTextColor($wParam, 0) ; чёрный
		Return $g_hBrushDrag
	EndIf
	Return __GUIDarkTheme_WM_CTLCOLOR($hWnd, $iMsg, $wParam, $lParam)
EndFunc   ;==>_OnEvent_WM_CTLCOLOREDIT


Func _OnEvent_GUI_EVENT_DROPPED()
	Local $iInput = _VideoInputOf(@GUI_DropId)
	Local $sDropFile = @GUI_DragFile
	If Not $iInput Or Not FileExists($sDropFile) Then Return

	If Not _IsValidVideoExtension($sDropFile) Then
		MsgBox($MB_ICONWARNING, $gc_sAppName, _
				Lang("Errors", "UnsupportedFormat", "Unsupported file format:") & " " & _GetFileExt($sDropFile) & _
				@CR & @CR & _
				Lang("Errors", "SupportedFormats", "Supported formats:") & " " & _GetVideoExtensionsFilter())
		Return
	EndIf

	_SetVideoFile($iInput, $sDropFile)
EndFunc   ;==>_OnEvent_GUI_EVENT_DROPPED


; Поле видео принимает файл всей рамкой: у Edit по краям поля по 7 px,
; и бросок на них иначе молча пропадал
Func _AcceptVideoDrop($iInput)
	GUICtrlSetState($iInput, $GUI_DROPACCEPTED)
	GUICtrlSetState(_SkinInputFrame($iInput), $GUI_DROPACCEPTED)
EndFunc   ;==>_AcceptVideoDrop


; Поле видео по ControlID самого поля или его рамки. 0 – контрол не из полей видео
Func _VideoInputOf($iCtrl)
	If $iCtrl = 0 Then Return 0
	If $iCtrl = $g_iInput1 Or $iCtrl = _SkinInputFrame($g_iInput1) Then Return $g_iInput1
	If $iCtrl = $g_iInput2 Or $iCtrl = _SkinInputFrame($g_iInput2) Then Return $g_iInput2
	Return 0
EndFunc   ;==>_VideoInputOf


; Поле потеряло фокус: пользователь закончил правку. Путь из поля видео
; проверяется, относительный ищется в папке текущего файла этого слота.
Func _OnEvent_WM_COMMAND($hWnd, $iMsg, $wParam, $lParam)
	#forceref $hWnd, $iMsg, $lParam
	If BitShift($wParam, 16) <> $EN_KILLFOCUS Then Return $GUI_RUNDEFMSG
	Local $iID = BitAND($wParam, 0xFFFF)

	If $iID = $g_iInputOffset Then
		_UpdateCommandField()
		Return $GUI_RUNDEFMSG
	EndIf
	If $iID <> $g_iInput1 And $iID <> $g_iInput2 Then Return $GUI_RUNDEFMSG

	Local $sPath = GUICtrlRead($iID)
	If $sPath = "" Then Return $GUI_RUNDEFMSG

	If Not _IsAbsolutePath($sPath) Then
		Local $sCurrent = ($iID = $g_iInput1) ? $g_sVideoFile1 : $g_sVideoFile2
		If Not FileExists($sCurrent) Then
			; Искать негде: поле сбрасываем, последний путь в ini не трогаем
			If $iID = $g_iInput1 Then
				$g_sVideoFile1 = ""
			Else
				$g_sVideoFile2 = ""
			EndIf
			GUICtrlSetData($iID, "")
			_UpdateFilesInfo()
			Return $GUI_RUNDEFMSG
		EndIf
		$sPath = _PathGetDir($sCurrent) & "\" & $sPath
	EndIf

	If Not FileExists($sPath) Or Not _IsValidVideoExtension($sPath) Then $sPath = ""
	_SetVideoFile($iID, $sPath)
	Return $GUI_RUNDEFMSG
EndFunc   ;==>_OnEvent_WM_COMMAND


; Сбрасывает слоты, чей файл исчез с диска. True – что-то сброшено,
; и вызывающий обязан обновить окно через _UpdateFilesInfo.
Func _ValidateVideoFiles()
	Local $bReset = False

	If $g_sVideoFile1 <> "" And Not FileExists($g_sVideoFile1) Then
		$g_sVideoFile1 = ""
		GUICtrlSetData($g_iInput1, "")
		$bReset = True
	EndIf

	If $g_sVideoFile2 <> "" And Not FileExists($g_sVideoFile2) Then
		$g_sVideoFile2 = ""
		GUICtrlSetData($g_iInput2, "")
		$bReset = True
	EndIf

	; Сдвиг считается для пары файлов, без одного из них он бессмыслен
	If $bReset Then
		GUICtrlSetData($g_iInputOffset, "")
		_TryFillCachedOffset() ; в авто-режиме поставит статус «не запускался»
	EndIf

	Return $bReset
EndFunc   ;==>_ValidateVideoFiles


; Пояснения под полями видео, доступность «Сравнить» и команда
Func _UpdateFilesInfo()
	_ValidateVideoFiles() ; выбранный файл мог исчезнуть с диска

	Local $bExists1 = FileExists($g_sVideoFile1), $bExists2 = FileExists($g_sVideoFile2)
	Local $aInfo1 = _GetVideoInfo($g_sVideoFile1), $aInfo2 = _GetVideoInfo($g_sVideoFile2)

	; Команда собирается, только когда у обоих файлов известно разрешение
	If Not $bExists1 Or Not $bExists2 Or $aInfo1[0] <= 0 Or $aInfo2[0] <= 0 Then
		GUICtrlSetData($g_iLabelInfo1, _FormatInfoLabel($aInfo1, $bExists1))
		GUICtrlSetData($g_iLabelInfo2, _FormatInfoLabel($aInfo2, $bExists2))
		_SkinBtnSetEnabled($g_iButtonCompare, False)
		GUICtrlSetData($g_iEditCommand, "")
		Return
	EndIf

	Local $aCropArgs = _CalculateCropArgs($aInfo1, $aInfo2)
	GUICtrlSetData($g_iLabelInfo1, _FormatInfoLabel($aInfo1, True, $aCropArgs[2]))
	GUICtrlSetData($g_iLabelInfo2, _FormatInfoLabel($aInfo2, True, $aCropArgs[3]))
	_SkinBtnSetEnabled($g_iButtonCompare, True)
	GUICtrlSetData($g_iEditCommand, _ComputeCommandFrom($aInfo1, $aInfo2, $aCropArgs))
EndFunc   ;==>_UpdateFilesInfo


; Пояснение под полем видео: «файл не выбран», «разрешение не определено»
; или разрешение, а при обрезке до $iCropH ещё и разница высоты.
Func _FormatInfoLabel($aInfo, $bExists, $iCropH = 0)
	If Not $bExists Then Return Lang("GUI", "FileNotSelected", "File not selected")
	If $aInfo[0] <= 0 Then Return Lang("Errors", "ResolutionUnknown", "Resolution not detected")

	Local $sText = Lang("Info", "Resolution", "Resolution") & " " & $aInfo[0] & "x" & $aInfo[1]
	If $iCropH > 0 And $iCropH <> $aInfo[1] Then _
		$sText &= ", " & Lang("Info", "HeightDiff", "height difference") & " " & ($aInfo[1] - $iCropH)
	Return $sText
EndFunc   ;==>_FormatInfoLabel


Func _UpdateCommandField()
	; Пропавший файл обнуляет не только команду, окно приводим к виду «файл не выбран»
	If _ValidateVideoFiles() Then
		_UpdateFilesInfo()
		Return
	EndIf
	GUICtrlSetData($g_iEditCommand, _ComputeCommand())
EndFunc   ;==>_UpdateCommandField


; Команда video-compare из текущего состояния окна. "" – нет файлов или разрешения.
Func _ComputeCommand()
	If Not FileExists($g_sVideoFile1) Or Not FileExists($g_sVideoFile2) Then Return ""
	Local $aVideo1Info = _GetVideoInfo($g_sVideoFile1)
	Local $aVideo2Info = _GetVideoInfo($g_sVideoFile2)
	If $aVideo1Info[0] <= 0 Or $aVideo2Info[0] <= 0 Then Return ""
	Local $aCropArgs = _CalculateCropArgs($aVideo1Info, $aVideo2Info)
	Return _ComputeCommandFrom($aVideo1Info, $aVideo2Info, $aCropArgs)
EndFunc   ;==>_ComputeCommand


; Команда из уже посчитанных разрешений и обрезки. Разрешения обязаны быть известны
Func _ComputeCommandFrom($aVideo1Info, $aVideo2Info, $aCropArgs)
	; Int("") = 0: пустое поле сдвига означает без сдвига
	Return _BuildVideoCompareCommand($aVideo1Info, $aVideo2Info, $aCropArgs, _IsVerticalMode(), _
			Int(GUICtrlRead($g_iInputOffset)))
EndFunc   ;==>_ComputeCommandFrom


; Разрешение файла [ширина, высота]. Неизвестное – пустые строки
Func _GetVideoInfo($sVideoPath)
	Local $sResolution = _GetVideoResolution($sVideoPath)
	Local $aInfo[2]
	Local $aSplit = StringSplit($sResolution, ',')
	If $aSplit[0] >= 2 Then
		$aInfo[0] = Int($aSplit[1])
		$aInfo[1] = Int($aSplit[2])
	EndIf
	Return $aInfo
EndFunc   ;==>_GetVideoInfo


; Обрезка лишней высоты снизу после приведения кадров к меньшей ширине.
; Возвращает [фильтр левого, фильтр правого, итоговая высота 1, итоговая высота 2].
Func _CalculateCropArgs($aVideo1Info, $aVideo2Info)
	Local $iW1 = $aVideo1Info[0], $iH1 = $aVideo1Info[1]
	Local $iW2 = $aVideo2Info[0], $iH2 = $aVideo2Info[1]
	Local $aCrop[4] = ["", "", $iH1, $iH2]

	; Без разрешения масштаб не посчитать: деление на ноль
	If $iW1 <= 0 Or $iW2 <= 0 Then Return $aCrop

	Local $iTargetW = _Min($iW1, $iW2)
	Local $nScale1 = $iTargetW / $iW1, $nScale2 = $iTargetW / $iW2
	Local $iScaledH1 = Round($iH1 * $nScale1), $iScaledH2 = Round($iH2 * $nScale2)
	If $iScaledH1 = $iScaledH2 Then Return $aCrop

	; Разницу высот переводим обратно в пиксели исходника, который обрезаем
	Local $iDiff
	If $iScaledH1 > $iScaledH2 Then
		$iDiff = Round(($iScaledH1 - $iScaledH2) / $nScale1)
		$aCrop[0] = "-l crop=iw:ih-" & $iDiff & " "
		$aCrop[2] = $iH1 - $iDiff
	Else
		$iDiff = Round(($iScaledH2 - $iScaledH1) / $nScale2)
		$aCrop[1] = "-r crop=iw:ih-" & $iDiff & " "
		$aCrop[3] = $iH2 - $iDiff
	EndIf
	Return $aCrop
EndFunc   ;==>_CalculateCropArgs


; Идентификатор файла без пути: имя и размер
Func _BuildFileId($sFile)
	Return _GetFileName($sFile) & "|" & FileGetSize($sFile)
EndFunc   ;==>_BuildFileId


; Ключ файла в кеше памяти: идентификатор и mtime
Func _BuildMemKey($sFile)
	Return _BuildFileId($sFile) & "|" & FileGetTime($sFile, $FT_MODIFIED, $FT_STRING)
EndFunc   ;==>_BuildMemKey


; Ключ пары файлов в кеше памяти
Func _BuildPairKey($sFile1, $sFile2)
	Return _BuildMemKey($sFile1) & "|" & _BuildMemKey($sFile2)
EndFunc   ;==>_BuildPairKey


; Ключ файла в [Info] кеша на диске. INI не читает обратно ключ с «=» внутри
; и с «;» или «[» в начале, поэтому они кодируются как в URL, а с ними и сам «%».
Func _BuildDiskKey($sFile)
	Local $sKey = StringReplace(StringReplace(_BuildFileId($sFile), "%", "%25"), "=", "%3D")
	If StringRegExp($sKey, "^[;\[]") Then $sKey = "%" & Hex(AscW($sKey), 2) & StringTrimLeft($sKey, 1)
	Return $sKey
EndFunc   ;==>_BuildDiskKey


; Запись файла в [Info] кеша: [индекс, разрешение]. Нет записи или mtime устарел – ["", ""]
Func _GetCacheInfo($sFile)
	Local $aEmpty[2] = ["", ""]
	Local $sDiskKey = _BuildDiskKey($sFile)
	Local $sCached = IniRead($gc_sPathCache, "Info", $sDiskKey, "")
	If $sCached = "" Then Return $aEmpty

	; Формат значения: индекс|mtime|разрешение (3 поля)
	Local $aParts = StringSplit($sCached, "|")
	If $aParts[0] < 3 Then Return $aEmpty

	If $aParts[2] <> FileGetTime($sFile, $FT_MODIFIED, $FT_STRING) Then Return $aEmpty

	Local $aResult[2] = [$aParts[1], $aParts[3]]
	Return $aResult
EndFunc   ;==>_GetCacheInfo


; Пишет файл в [Info] кеша и возвращает его индекс: прежний или следующий из [Meta]
Func _SaveCacheInfo($sFile, $sResolution)
	Local $sDiskKey = _BuildDiskKey($sFile)
	Local $sCached = IniRead($gc_sPathCache, "Info", $sDiskKey, "")
	Local $iIdx
	If $sCached <> "" Then
		$iIdx = Int(StringSplit($sCached, "|")[1])
	Else
		$iIdx = Int(IniRead($gc_sPathCache, "Meta", "NextId", "1"))
		IniWrite($gc_sPathCache, "Meta", "NextId", $iIdx + 1)
	EndIf

	IniWrite($gc_sPathCache, "Info", $sDiskKey, _
			$iIdx & "|" & FileGetTime($sFile, $FT_MODIFIED, $FT_STRING) & "|" & $sResolution)
	Return $iIdx
EndFunc   ;==>_SaveCacheInfo


; Значение кеша sync – маркер неудачи. Сравнение строгое: при «=» число 0,
; то есть найденный нулевой сдвиг, равно любой строке и выглядело бы маркером.
Func _IsSyncMarker($vValue)
	Return ($vValue == "TIMEOUT" Or $vValue == "NOMATCH" Or $vValue == "ERROR")
EndFunc   ;==>_IsSyncMarker


; Сдвиг из кеша: память, затем диск, для прямой и обратной пары.
; Возвращает [$bFound, $iOffset, $sStatus]; не найдено – [False, 0, "ERROR"].
; Найденное на диске или по обратной паре кладётся в память под прямой ключ.
Func _LookupSyncCache($sFile1, $sFile2)
	Local $aNotFound[3] = [False, 0, "ERROR"]
	Local $sMemKey = _BuildPairKey($sFile1, $sFile2)
	If MapExists($g_oCache, $sMemKey) Then Return __DecodeSyncVal($g_oCache[$sMemKey], False)

	Local $aDec, $sVal
	Local $sMemKeyRev = _BuildPairKey($sFile2, $sFile1)
	If MapExists($g_oCache, $sMemKeyRev) Then
		$aDec = __DecodeSyncVal($g_oCache[$sMemKeyRev], True)
	Else
		; На диске пары лежат по индексам файлов из [Info]
		Local $aInfo1 = _GetCacheInfo($sFile1), $aInfo2 = _GetCacheInfo($sFile2)
		If $aInfo1[0] = "" Or $aInfo2[0] = "" Then Return $aNotFound

		$sVal = IniRead($gc_sPathCache, "Sync", $aInfo1[0] & "|" & $aInfo2[0], "")
		If $sVal <> "" Then
			$aDec = __DecodeSyncVal($sVal, False)
		Else
			$sVal = IniRead($gc_sPathCache, "Sync", $aInfo2[0] & "|" & $aInfo1[0], "")
			If $sVal = "" Then Return $aNotFound
			$aDec = __DecodeSyncVal($sVal, True)
		EndIf
	EndIf

	$g_oCache[$sMemKey] = ($aDec[2] = "OK") ? $aDec[1] : $aDec[2]
	Return $aDec
EndFunc   ;==>_LookupSyncCache


; Значение кеша sync в [True, $iOffset, $sStatus]: маркер неудачи даёт нулевой сдвиг,
; число – статус OK. $bReverse меняет знак сдвига для обратной пары файлов.
Func __DecodeSyncVal($vVal, $bReverse)
	Local $aOut[3] = [True, 0, "ERROR"]
	If _IsSyncMarker($vVal) Then
		$aOut[2] = $vVal
	Else
		$aOut[1] = $bReverse ? -Int($vVal) : Int($vVal)
		$aOut[2] = "OK"
	EndIf
	Return $aOut
EndFunc   ;==>__DecodeSyncVal


; Результат sync в кеш памяти и диска. $vValue – сдвиг числом или маркер неудачи
Func _SaveSyncCache($sFile1, $sFile2, $vValue)
	$g_oCache[_BuildPairKey($sFile1, $sFile2)] = $vValue
	$g_oCache[_BuildPairKey($sFile2, $sFile1)] = _IsSyncMarker($vValue) ? $vValue : -$vValue

	; Файл без записи в [Info] получает её вместе с разрешением: пустое поле
	; заставило бы потом гонять ffmpeg повторно
	Local $aInfo1 = _GetCacheInfo($sFile1), $aInfo2 = _GetCacheInfo($sFile2)
	Local $iIdx1 = ($aInfo1[0] <> "") ? Int($aInfo1[0]) : _SaveCacheInfo($sFile1, _GetVideoResolution($sFile1))
	Local $iIdx2 = ($aInfo2[0] <> "") ? Int($aInfo2[0]) : _SaveCacheInfo($sFile2, _GetVideoResolution($sFile2))
	IniWrite($gc_sPathCache, "Sync", $iIdx1 & "|" & $iIdx2, $vValue)
EndFunc   ;==>_SaveSyncCache


; Сдвиг пары файлов: [$sStatus, $iOffset]. Любая запись в кеше, в том числе неудача,
; возвращается без повторного поиска. Исход поиска кешируется всегда.
Func _GetSyncOffset($sFile1, $sFile2)
	Local $aCached = _LookupSyncCache($sFile1, $sFile2)
	If $aCached[0] Then
		Local $aHit[2] = [$aCached[2], $aCached[1]]
		Return $aHit
	EndIf

	Local $aRun = _RunSyncWithTimeout($sFile1, $sFile2, $g_iSyncTimeoutSec)
	If $aRun[0] = "NOMATCH" Then
		; Разные озвучки роняют корреляцию звука, а смены сцен от дорожки не зависят
		Local $sAudioOutput = $g_sLastSyncOutput
		$aRun = _RunSyncWithTimeout($sFile1, $sFile2, $g_iSyncTimeoutSec, "video")
		; В деталях оба прогона
		$g_sLastSyncOutput = $sAudioOutput & @CRLF & "--- video fallback ---" & @CRLF & $g_sLastSyncOutput
	ElseIf $aRun[0] = "OK" And Abs($aRun[1]) > $g_iSyncSuspectMs Then
		$aRun = _VerifyOffsetByVideo($sFile1, $sFile2, $aRun)
	EndIf

	_SaveSyncCache($sFile1, $sFile2, ($aRun[0] = "OK") ? $aRun[1] : $aRun[0])
	Return $aRun
EndFunc   ;==>_GetSyncOffset


; Подозрительно большой сдвиг по звуку перепроверяется по видео.
; Совпал в пределах допуска – исходный результат, иначе ["NOMATCH", 0].
Func _VerifyOffsetByVideo($sFile1, $sFile2, $aPrimary)
	Local $aNoMatch[2] = ["NOMATCH", 0]
	Local $aVideo = _RunSyncWithTimeout($sFile1, $sFile2, $g_iSyncTimeoutSec, "video")
	If $aVideo[0] <> "OK" Then Return $aNoMatch
	If Abs($aVideo[1] - $aPrimary[1]) > $g_iSyncVerifyTolMs Then Return $aNoMatch
	Return $aPrimary
EndFunc   ;==>_VerifyOffsetByVideo


; Запускает Sync.exe и ждёт результат не дольше $iTimeoutSec. Возвращает [$sStatus, $iOffset]:
; OK – exit 0 и число в stdout, TIMEOUT – процесс убит по таймауту,
; NOMATCH – ненулевой exit, ERROR – exit 0 без числа.
Func _RunSyncWithTimeout($sFile1, $sFile2, $iTimeoutSec, $sMethod = "audio")
	Local $sCmdLine = '"' & $g_sPathSync & '" sync --method ' & $sMethod & _
			' --v1 "' & $sFile1 & '" --v2 "' & $sFile2 & '" --skip ' & $g_iSyncSkipSec & _
			' --ffmpeg "' & $g_sPathFFmpeg & '"'
	$g_sLastSyncCmd = $sCmdLine ; для деталей по клику на статус
	$g_sLastSyncOutput = ""

	Local $aResult[2] = ["ERROR", 0]
	Local $iPid = Run($sCmdLine, "", @SW_HIDE, $STDERR_CHILD + $STDOUT_CHILD)
	If @error Then Return $aResult
	; Handle держим ради кода выхода: Run его не отдаёт
	Local $hProcess = _WinAPI_OpenProcess($STANDARD_RIGHTS_SYNCHRONIZE + $PROCESS_QUERY_INFORMATION, False, $iPid)
	Local $hTimer = TimerInit()
	Local $sStdout = "", $sStderr = ""

	While ProcessExists($iPid)
		If TimerDiff($hTimer) >= $iTimeoutSec * 1000 Then
			ProcessClose($iPid)
			ConsoleWrite("Sync.exe: таймаут (" & $iTimeoutSec & " сек)" & @CRLF)
			If $hProcess Then _WinAPI_CloseHandle($hProcess)
			$g_sLastSyncOutput = _ComposeSyncOutput($sStdout, $sStderr)
			$aResult[0] = "TIMEOUT"
			Return $aResult
		EndIf

		; Потоки читаем на ходу, иначе заполненный pipe остановит процесс
		$sStdout &= StdoutRead($iPid)
		$sStderr &= StderrRead($iPid)
		Sleep(100)
	WEnd

	$sStdout &= _DrainPipe($iPid, False)
	$sStderr &= _DrainPipe($iPid, True)
	$g_sLastSyncOutput = _ComposeSyncOutput($sStdout, $sStderr)

	Local $iExitCode = 0
	If $hProcess Then
		$iExitCode = _WinAPI_GetExitCodeProcess($hProcess)
		_WinAPI_CloseHandle($hProcess)
	EndIf

	Local $sTrim = StringStripWS($sStdout, $STR_STRIPLEADING + $STR_STRIPTRAILING)
	If $iExitCode <> 0 Then
		$aResult[0] = "NOMATCH"
	ElseIf StringRegExp($sTrim, "^-?\d+$") Then
		$aResult[0] = "OK"
		$aResult[1] = Int($sTrim)
	EndIf
	Return $aResult
EndFunc   ;==>_RunSyncWithTimeout


; Stdout и stderr одним текстом с заголовками – для деталей по клику на статус
Func _ComposeSyncOutput($sStdout, $sStderr)
	Local $sOut = "", $iStrip = $STR_STRIPLEADING + $STR_STRIPTRAILING
	$sStdout = StringStripWS($sStdout, $iStrip)
	$sStderr = StringStripWS($sStderr, $iStrip)
	If $sStdout <> "" Then $sOut = "[stdout]" & @CRLF & $sStdout
	If $sStderr <> "" Then $sOut &= ($sOut = "" ? "" : @CRLF & @CRLF) & "[stderr]" & @CRLF & $sStderr
	Return $sOut
EndFunc   ;==>_ComposeSyncOutput


; Читает поток процесса до EOF. $bStderr – stderr, иначе stdout.
; StdoutRead не блокирует: пока процесс жив, без паузы на пустом чтении
; цикл занимал бы ядро процессора целиком.
Func _DrainPipe($iPid, $bStderr)
	Local $sOut = "", $sChunk
	While 1
		$sChunk = $bStderr ? StderrRead($iPid) : StdoutRead($iPid)
		If @error Then ExitLoop
		$sOut &= $sChunk
		If $sChunk = "" Then Sleep(10)
	WEnd
	Return $sOut
EndFunc   ;==>_DrainPipe


; Поле сдвига сбрасывается, только если пара видео реально сменилась:
; значение той же пары переживает повторный выбор файла и потерю фокуса.
Func _ResetOffsetForNewPair($sPrevFile1, $sPrevFile2)
	If $sPrevFile1 = $g_sVideoFile1 And $sPrevFile2 = $g_sVideoFile2 Then Return
	GUICtrlSetData($g_iInputOffset, "")
	_TryFillCachedOffset()
EndFunc   ;==>_ResetOffsetForNewPair


Func _TryFillCachedOffset()
	If Not _IsAutoOffset() Then Return
	If Not FileExists($g_sVideoFile1) Or Not FileExists($g_sVideoFile2) Then
		_SetSyncStatus("NOTRUN")
		Return
	EndIf

	Local $aCached = _LookupSyncCache($g_sVideoFile1, $g_sVideoFile2)
	If Not $aCached[0] Then
		_SetSyncStatus("NOTRUN")
		Return
	EndIf

	GUICtrlSetData($g_iInputOffset, ($aCached[2] = "OK") ? $aCached[1] : "")
	_SetSyncStatus($aCached[2], $aCached[1])
EndFunc   ;==>_TryFillCachedOffset


; Статус поиска сдвига под полем: текст и цвет. Статусы – см. $g_sLastSyncStatus,
; %d в тексте OK заменяется сдвигом.
Func _SetSyncStatus($sStatus, $iOffset = 0)
	_SyncSpinnerStop()
	$g_sLastSyncStatus = $sStatus
	$g_iLastSyncOffset = $iOffset

	If $sStatus = "" Then
		GUICtrlSetData($g_iLabelSyncStatus, "")
		Return
	EndIf

	; Пока результата нет, статус серый
	Local $iColor = $g_iClrInfo, $sTextKey = ""
	Switch $sStatus
		Case "NOTRUN"
			$sTextKey = "StatusNotRun"
		Case "MANUAL"
			$sTextKey = "StatusManual"
		Case "WORKING"
			$sTextKey = "StatusWorking"
			_SyncSpinnerStart() ; текст вместе с кадром рисует _SyncSpinnerRender
		Case "OK"
			$iColor = $g_bDarkMode ? $gc_iClrOkDark : $gc_iClrOkLight
			$sTextKey = "StatusOk"
		Case "NOMATCH"
			$iColor = $g_bDarkMode ? $gc_iClrErrDark : $gc_iClrErrLight
			$sTextKey = "StatusNoMatch"
		Case "TIMEOUT"
			$iColor = $g_bDarkMode ? $gc_iClrWarnDark : $gc_iClrWarnLight
			$sTextKey = "StatusTimeout"
		Case "ERROR"
			$iColor = $g_bDarkMode ? $gc_iClrErrDark : $gc_iClrErrLight
			$sTextKey = "StatusError"
		Case Else
			Return
	EndSwitch

	If $sStatus <> "WORKING" Then
		GUICtrlSetData($g_iLabelSyncStatus, StringReplace(Lang("Sync", $sTextKey, $sStatus), "%d", $iOffset))
	EndIf
	GUICtrlSetColor($g_iLabelSyncStatus, $iColor)
	; GUICtrlSetColor инвалидирует родителя, а тот с WS_CLIPCHILDREN подпись обходит
	_WinAPI_InvalidateRect(GUICtrlGetHandle($g_iLabelSyncStatus))
EndFunc   ;==>_SetSyncStatus


Func _SyncSpinnerStart()
	$g_iSpinnerIdx = 0
	_SyncSpinnerRender()
	AdlibRegister("_SyncSpinnerTick", 120)
EndFunc   ;==>_SyncSpinnerStart


Func _SyncSpinnerStop()
	AdlibUnRegister("_SyncSpinnerTick")
EndFunc   ;==>_SyncSpinnerStop


Func _SyncSpinnerTick()
	$g_iSpinnerIdx = Mod($g_iSpinnerIdx + 1, UBound($gc_aSpinnerFrames))
	_SyncSpinnerRender()
EndFunc   ;==>_SyncSpinnerTick


Func _SyncSpinnerRender()
	GUICtrlSetData($g_iLabelSyncStatus, $gc_aSpinnerFrames[$g_iSpinnerIdx] & " " & _
			Lang("Sync", "StatusWorking", "searching…"))
EndFunc   ;==>_SyncSpinnerRender


; Клик по статусу неудачи: MsgBox с командой и выводом Sync.exe
Func _OnEvent_LabelSyncStatusClick()
	If Not _IsSyncMarker($g_sLastSyncStatus) Then Return

	Local $sBody = ""
	If $g_sLastSyncCmd <> "" Then $sBody &= "[command]" & @CRLF & $g_sLastSyncCmd & @CRLF & @CRLF
	If $g_sLastSyncOutput <> "" Then $sBody &= $g_sLastSyncOutput
	If $sBody = "" Then $sBody = Lang("Sync", "NoDetails", "No details available.")

	; Длинный текст MsgBox обрезает: полный вывод уходит в консоль
	Local Const $iMaxLen = 4000
	If StringLen($sBody) > $iMaxLen Then
		ConsoleWrite($sBody & @CRLF)
		$sBody = StringLeft($sBody, $iMaxLen) & @CRLF & @CRLF & "... (truncated, see console)"
	EndIf

	MsgBox($MB_ICONINFORMATION, $gc_sAppName & " – sync details", $sBody)
EndFunc   ;==>_OnEvent_LabelSyncStatusClick


Func _BuildVideoCompareCommand($aVideo1Info, $aVideo2Info, $aCropArgs, $bIsVertical, $iOffsetMs = 0)
	Local $sCmdLine = '"' & $g_sPathVideoCompare & '"'

	; native: video-compare сам кладёт кадры 1:1 по центру холста, обрезка не нужна.
	; Ключ --conversion-fit есть со сборки 20260828, старая на нём оборвала бы разбор
	; аргументов, поэтому для неё остаётся обрезка.
	Local $bNative = ($g_sFit = "native")
	If $bNative Then
		Local $iVcDate = _VcBuildDate()
		$bNative = ($iVcDate = 0 Or $iVcDate >= 20260828)
	EndIf

	If _ShouldUseFullscreen($aVideo1Info, $aVideo2Info, $aCropArgs, $bIsVertical, $bNative) Then $sCmdLine &= " -W"
	If $bIsVertical Then $sCmdLine &= " -m vstack"
	If $bNative Then $sCmdLine &= " --conversion-fit native"
	If $iOffsetMs <> 0 Then $sCmdLine &= " -t " & $iOffsetMs / 1000 ; ключ ждёт секунды

	Local $sCrop = $bNative ? "" : $aCropArgs[0] & $aCropArgs[1]
	Return $sCmdLine & " " & $sCrop & '"' & $g_sVideoFile1 & '" "' & $g_sVideoFile2 & '"'
EndFunc   ;==>_BuildVideoCompareCommand


; Холст сравнения больше рабочей области экрана – video-compare нужен ключ -W
Func _ShouldUseFullscreen($aVideo1Info, $aVideo2Info, $aCropArgs, $bIsVertical, $bNative)
	Local $tWork = _WinAPI_GetWorkArea()
	Local $iDesktopWidth = DllStructGetData($tWork, "Right") - DllStructGetData($tWork, "Left")
	Local $iDesktopHeight = DllStructGetData($tWork, "Bottom") - DllStructGetData($tWork, "Top")

	Local $iMaxWidth = _Max($aVideo1Info[0], $aVideo2Info[0])

	; native: высоту задаёт холст, то есть большая из исходных; crop: уже выровненные.
	; Вертикальное расположение высоты складывает.
	Local $iMaxHeight = $bNative ? _Max($aVideo1Info[1], $aVideo2Info[1]) _
			: _Max($aCropArgs[2], $aCropArgs[3])
	If $bIsVertical Then $iMaxHeight *= 2

	Return ($iMaxWidth > $iDesktopWidth Or $iMaxHeight > $iDesktopHeight)
EndFunc   ;==>_ShouldUseFullscreen


; Запускает video-compare без консоли и ждёт выхода, собирая stdout и stderr.
; Ненулевой код выхода – MsgBox с извлечённой строкой ошибки.
; CreateProcessW вместо Run ради CREATE_NO_WINDOW: @SW_HIDE из Run ушёл бы в STARTUPINFO,
; и первый ShowWindow спрятал бы само окно сравнения.
Func _RunVideoCompare($sCmdLine)
	ConsoleWrite($sCmdLine & @CRLF)

	; Pipe с наследуемыми handle. Читающий конец процессу не отдаём, иначе EOF не придёт
	Local $tSA = DllStructCreate($tagSECURITY_ATTRIBUTES)
	DllStructSetData($tSA, "Length", DllStructGetSize($tSA))
	DllStructSetData($tSA, "InheritHandle", True)
	Local $aPipe = DllCall("kernel32.dll", "bool", "CreatePipe", _
			"handle*", 0,    _
			"handle*", 0,    _
			"struct*", $tSA, _
			"dword",   0)
	If @error Or Not $aPipe[0] Then Return
	Local $hRead = $aPipe[1], $hWrite = $aPipe[2]
	_WinAPI_SetHandleInformation($hRead, $HANDLE_FLAG_INHERIT, 0)

	Local $tSI = DllStructCreate($tagSTARTUPINFO)
	DllStructSetData($tSI, "Size", DllStructGetSize($tSI))
	DllStructSetData($tSI, "Flags", $STARTF_USESTDHANDLES)
	DllStructSetData($tSI, "StdOutput", $hWrite)
	DllStructSetData($tSI, "StdError", $hWrite)
	Local $tPI = DllStructCreate($tagPROCESS_INFORMATION)
	; CreateProcessW пишет в строку команды, поэтому она в своём буфере
	Local $tCmd = DllStructCreate("wchar[" & StringLen($sCmdLine) + 1 & "]")
	DllStructSetData($tCmd, 1, $sCmdLine)

	Local $aCP = DllCall("kernel32.dll", "bool", "CreateProcessW", _
			"ptr",     0,                 _
			"struct*", $tCmd,             _
			"ptr",     0, "ptr", 0,       _
			"bool",    True,              _
			"dword",   $CREATE_NO_WINDOW, _
			"ptr",     0, "ptr", 0,       _
			"struct*", $tSI,              _
			"struct*", $tPI)
	Local $bStarted = Not @error And $aCP[0] ; до CloseHandle: он сбросит @error

	; Пишущий конец остался только у процесса: закроет его – придёт EOF
	_WinAPI_CloseHandle($hWrite)
	If Not $bStarted Then
		_WinAPI_CloseHandle($hRead)
		Return
	EndIf

	; Неблокирующее чтение: в паузах Sleep окно продолжает перерисовываться
	Local $hProcess = DllStructGetData($tPI, "hProcess")
	Local $tBuf = DllStructCreate("char[4096]")
	Local $tAvail = DllStructCreate("dword")
	Local $sOutput = "", $iRead = 0, $bExited = False
	While 1
		Local $aPeek = DllCall("kernel32.dll", "bool", "PeekNamedPipe", _
				"handle",  $hRead,  _
				"ptr",     0,       _
				"dword",   0,       _
				"ptr",     0,       _
				"struct*", $tAvail, _
				"ptr",     0)
		If @error Or Not $aPeek[0] Then ExitLoop ; pipe закрыт и пуст

		If DllStructGetData($tAvail, 1) = 0 Then
			; Pipe может держать открытым потомок процесса: после выхода самого
			; процесса хватает одного круга, чтобы дочитать успевшее прийти
			If $bExited Then ExitLoop
			$bExited = (_WinAPI_WaitForSingleObject($hProcess, 0) = 0) ; WAIT_OBJECT_0
			If Not $bExited Then Sleep(30)
			ContinueLoop
		EndIf

		If Not _WinAPI_ReadFile($hRead, $tBuf, 4096, $iRead) Or $iRead = 0 Then ExitLoop
		; Буфер без нулевого терминатора и с хвостом прошлого чтения
		$sOutput &= StringLeft(DllStructGetData($tBuf, 1), $iRead)
	WEnd

	; Pipe закрывается чуть раньше, чем процесс становится signaled.
	; Без ожидания код выхода мог прийти как STILL_ACTIVE (259) и дать ложную ошибку.
	_WinAPI_WaitForSingleObject($hProcess, 5000)
	Local $iExitCode = _WinAPI_GetExitCodeProcess($hProcess)
	_WinAPI_CloseHandle($hRead)
	_WinAPI_CloseHandle($hProcess)
	_WinAPI_CloseHandle(DllStructGetData($tPI, "hThread"))

	; Вывод консольный, в OEM (CP866 на русской Windows). CP_OEMCP = 1
	If $sOutput <> "" Then $sOutput = _WinAPI_MultiByteToWideChar($sOutput, 1, 0, True)

	; Код выхода к знаковому: -1 читается лучше, чем 4294967295
	If $iExitCode > 0x7FFFFFFF Then $iExitCode -= 0x100000000
	ConsoleWrite($sOutput & @CRLF & "[exit=" & $iExitCode & "]" & @CRLF)
	If $iExitCode <> 0 Then _ShowVideoCompareError($sOutput, $iExitCode)
EndFunc   ;==>_RunVideoCompare


Func _ShowVideoCompareError($sOutput, $iExitCode)
	Local $sMsg = _ExtractErrorLines($sOutput)
	If $sMsg = "" Then $sMsg = StringStripWS($sOutput, 3)
	If $sMsg = "" Then $sMsg = Lang("Errors", "VideoCompareNoOutput", "No error message in output.")

	; Длинный вывод обрезаем с начала: ошибка обычно в конце
	Local Const $iMaxLen = 1500
	If StringLen($sMsg) > $iMaxLen Then $sMsg = "..." & StringRight($sMsg, $iMaxLen)

	MsgBox($MB_ICONERROR, $gc_sAppName, Lang("Errors", "VideoCompareFailed", "Video-compare exited with an error") & _
			" (exit " & $iExitCode & ")" & @CRLF & @CRLF & $sMsg)
EndFunc   ;==>_ShowVideoCompareError


; Строки вывода с маркерами ошибок video-compare и FFmpeg
Func _ExtractErrorLines($sOutput)
	; Одиночный CR тоже разрыв строки: им FFmpeg перезаписывает строку прогресса
	Local $aLines = StringSplit(StringRegExpReplace($sOutput, "\r\n?", @LF), @LF)
	Local $aMarkers[5] = ["Error:", "Exception", "terminate called", "Assertion", "Fatal"]
	Local $sResult = ""
	For $i = 1 To $aLines[0]
		For $sMarker In $aMarkers
			If StringInStr($aLines[$i], $sMarker) Then
				$sResult &= $aLines[$i] & @CRLF
				ExitLoop
			EndIf
		Next
	Next
	Return StringStripWS($sResult, $STR_STRIPLEADING + $STR_STRIPTRAILING)
EndFunc   ;==>_ExtractErrorLines


; Разрешение "W,H" из кеша памяти, кеша на диске или ffmpeg. Не определилось – ""
Func _GetVideoResolution($sVideoPath)
	If $sVideoPath = "" Or Not FileExists($sVideoPath) Then Return ""

	Local $sMemKey = _BuildMemKey($sVideoPath)
	If MapExists($g_oCache, $sMemKey) Then Return $g_oCache[$sMemKey]

	Local $aInfo = _GetCacheInfo($sVideoPath)
	If $aInfo[1] <> "" Then
		$g_oCache[$sMemKey] = $aInfo[1]
		Return $aInfo[1]
	EndIf

	; ffmpeg -i без выходного файла печатает свойства потоков в stderr и выходит
	Local $sRaw = _RunToolReadStderr('"' & $g_sPathFFmpeg & '" -hide_banner -i "' & $sVideoPath & '"')
	Local $aMatch = StringRegExp($sRaw, "Video:\s.*?,\s(\d+)x(\d+)", $STR_REGEXPARRAYMATCH)
	Local $sOutput = IsArray($aMatch) ? $aMatch[0] & "," & $aMatch[1] : ""
	ConsoleWrite($sVideoPath & " -> " & $sOutput & @CRLF)

	If $sOutput <> "" Then
		$g_oCache[$sMemKey] = $sOutput
		_SaveCacheInfo($sVideoPath, $sOutput)
	EndIf
	Return $sOutput
EndFunc   ;==>_GetVideoResolution


; Запускает инструмент и возвращает его stderr целиком: ffmpeg пишет метаданные туда.
Func _RunToolReadStderr($sCmdLine)
	Local $iPid = Run($sCmdLine, "", @SW_HIDE, $STDERR_CHILD)
	If @error Then Return ""
	Return StringStripWS(_DrainPipe($iPid, True), $STR_STRIPLEADING + $STR_STRIPTRAILING)
EndFunc   ;==>_RunToolReadStderr


; Короткие пояснения под переключателями «Режима», полный текст – в подсказке
Func _UpdateModeDesc()
	Local $bVertical = _IsVerticalMode()
	GUICtrlSetData($g_iLabelModeDesc, $bVertical _
			? Lang("GUI", "CompareVertical", "Vertical compare") _
			: Lang("GUI", "CompareDirect", "Direct compare"))
	_SetTip($g_iLabelModeDesc, _CompareTip($bVertical))

	GUICtrlSetData($g_iLabelFitDesc, _FitDesc($g_sFit))
	_SetTip($g_iLabelFitDesc, _FitTip($g_sFit))
EndFunc   ;==>_UpdateModeDesc


; Подсказки сегментов «Режима»: у каждого сегмента текст своего режима
Func _UpdateModeTips()
	_SegSetTips($g_iSegCompare, _CompareTip(False) & "|" & _CompareTip(True), "_SetTip")
	_SegSetTips($g_iSegFit, _FitTip("native") & "|" & _FitTip("crop"), "_SetTip")
EndFunc   ;==>_UpdateModeTips


Func _CompareTip($bVertical)
	If $bVertical Then Return Lang("GUI", "CompareVerticalTip", "Frames sit one under another")
	Return Lang("GUI", "CompareDirectTip", "Frames side by side, the boundary follows the mouse")
EndFunc   ;==>_CompareTip


; $sFit – значение из ini: "native" | "crop"
Func _FitDesc($sFit)
	If $sFit = "crop" Then Return Lang("GUI", "FitCropDesc", "Overlay with cropping")
	Return Lang("GUI", "FitNativeDesc", "Native 1:1 overlay")
EndFunc   ;==>_FitDesc


Func _FitTip($sFit)
	If $sFit = "crop" Then Return Lang("GUI", "FitCropTip", _
			"Frames are scaled to a shared size, extra height is cropped")
	Return Lang("GUI", "FitNativeTip", _
			"Frames are overlaid unscaled, the smaller centred on the larger")
EndFunc   ;==>_FitTip


Func _IsVerticalMode()
	Return (_SegGetSel($g_iSegCompare) = 1)
EndFunc   ;==>_IsVerticalMode


Func _IsAutoOffset()
	Return (_SegGetSel($g_iSegOffset) = 0)
EndFunc   ;==>_IsAutoOffset


; Иконка «Сравнить» повторяет выбранное расположение кадров
Func _UpdateCompareButtonIcon()
	_SkinBtnSetIcon($g_iButtonCompare, _IsVerticalMode() ? "CompareVstack" : "CompareDirect")
EndFunc   ;==>_UpdateCompareButtonIcon


; Снимает drag-подсветку после броска файла
Func _RestoreControlsStyle()
	AdlibUnRegister("_RestoreControlsStyle")
	Local $hPrev = $g_hDragHiCtrl
	$g_hDragHiCtrl = 0
	If $hPrev Then _WinAPI_InvalidateRect($hPrev)
	_UpdateFilesInfo()
EndFunc   ;==>_RestoreControlsStyle


Func _GetVideoExtensionsFilter()
	Local $sExtensions = ""
	For $sExt In $gc_aSupportedExtensions
		$sExtensions &= "*." & $sExt & ";"
	Next
	Return Lang("Filter", "Video", "Video") & " (" & StringTrimRight($sExtensions, 1) & ")"
EndFunc   ;==>_GetVideoExtensionsFilter


; Папка файла без завершающего разделителя, понимает оба слэша. Без разделителя – ""
Func _PathGetDir($sPath)
	If Not StringRegExp($sPath, '[\\/]') Then Return ""
	Return StringRegExpReplace($sPath, '[\\/][^\\/]*$', '')
EndFunc   ;==>_PathGetDir


Func _GetFileName($sPath)
	Return StringRegExpReplace($sPath, '^.*[\\/]', '')
EndFunc   ;==>_GetFileName


; Расширение файла в нижнем регистре, без точки. Файл без расширения даёт имя целиком.
Func _GetFileExt($sPath)
	Return StringLower(StringRegExpReplace(_GetFileName($sPath), '^.*\.', ''))
EndFunc   ;==>_GetFileExt


; Выделение снимается, каретка в начало: длинное имя файла видно с начала
Func _ResetInputCaret($iCtrlID)
	GUICtrlSendMsg($iCtrlID, $EM_SETSEL, 0, 0)
EndFunc   ;==>_ResetInputCaret


; Имя файла в поле видео. Пустой путь очищает поле
Func _ShowVideoFile($iInput, $sFile)
	GUICtrlSetData($iInput, _GetFileName($sFile))
	_ResetInputCaret($iInput)
EndFunc   ;==>_ShowVideoFile


Func _IsValidVideoExtension($sFilePath)
	Local $sExt = _GetFileExt($sFilePath)
	For $sSupportedExt In $gc_aSupportedExtensions
		If $sExt = $sSupportedExt Then Return True
	Next
	Return False
EndFunc   ;==>_IsValidVideoExtension


; Файл в слот поля $iInputID: поле, ini, сдвиг для новой пары, пояснения и команда
Func _SetVideoFile($iInputID, $sFilePath)
	Local $sPrevFile1 = $g_sVideoFile1, $sPrevFile2 = $g_sVideoFile2

	If $iInputID = $g_iInput1 Then
		$g_sVideoFile1 = $sFilePath
		IniWrite($gc_sPathIni, "LastDirs", "Video1", $sFilePath)
	Else
		$g_sVideoFile2 = $sFilePath
		IniWrite($gc_sPathIni, "LastDirs", "Video2", $sFilePath)
	EndIf
	_ShowVideoFile($iInputID, $sFilePath)

	_ResetOffsetForNewPair($sPrevFile1, $sPrevFile2)
	_UpdateFilesInfo()
EndFunc   ;==>_SetVideoFile


; Ini с UTF-16 LE BOM (Unicode-пути) и значениями по умолчанию. Пути инструментов –
; раскладка сборки. Пустые Language и Theme при первом запуске берутся из системы.
Func _EnsureIniDefaults()
	If FileExists($gc_sPathIni) Then Return
	Local $sContent = "[LastDirs]" & @CRLF & _
			"Video1=" & @CRLF & _
			"Video2=" & @CRLF & _
			"[Tools]" & @CRLF & _
			"VideoCompare=video-compare.exe" & @CRLF & _
			"Sync=Sync\Sync.exe" & @CRLF & _
			"FFmpeg=ffmpeg.exe" & @CRLF & _
			"[Settings]" & @CRLF & _
			"Language=" & @CRLF & _
			"Theme=" & @CRLF & _
			"Fit=native" & @CRLF & _
			"Hotkeys=0" & @CRLF & _
			"SyncSkipSec=300" & @CRLF & _
			"SyncTimeoutSec=60" & @CRLF & _
			"SyncSuspectMs=10000" & @CRLF & _
			"SyncVerifyTolMs=500" & @CRLF & _
			"VcVersion=" & @CRLF & _
			"UpdateCheck=1" & @CRLF & _
			"LastUpdateCheck=" & @CRLF
	_EnsureUtf16File($gc_sPathIni, $sContent)
EndFunc   ;==>_EnsureIniDefaults


; Создаёт файл с UTF-16 LE BOM, если его ещё нет
Func _EnsureUtf16File($sFilePath, $sContent = "")
	If FileExists($sFilePath) Then Return
	Local $hFile = FileOpen($sFilePath, $FO_OVERWRITE + $FO_UTF16_LE)
	If $hFile = -1 Then Return
	If $sContent <> "" Then FileWrite($hFile, $sContent)
	FileClose($hFile)
EndFunc   ;==>_EnsureUtf16File


; Диск с корнем, UNC или путь от корня
Func _IsAbsolutePath($sPath)
	Return StringRegExp($sPath, "^(?:[A-Za-z]:\\|\\\\|/)") = 1
EndFunc   ;==>_IsAbsolutePath


; Относительный путь из ini считается от папки скрипта. Пустой остаётся пустым,
; иначе превратился бы в путь к самой папке.
Func _NormalizePath($sValue)
	If $sValue = "" Or _IsAbsolutePath($sValue) Then Return $sValue
	Return @ScriptDir & "\" & $sValue
EndFunc   ;==>_NormalizePath


; Нет инструмента – сообщение, открытие ini и выход
Func _CheckToolExists($sToolPath, $sToolName)
	If FileExists($sToolPath) Then Return
	MsgBox($MB_ICONWARNING, $gc_sAppName, _
			StringReplace(Lang("Errors", "ToolNotFound", '"%TOOL%" not found.'), "%TOOL%", $sToolName) & @CR & @CR & _
			Lang("Errors", "SettingsWillOpen", "The settings file will now be opened."))
	ShellExecute($gc_sPathIni)
	Exit
EndFunc   ;==>_CheckToolExists


Func _ResolveToolPaths()
	$g_sPathVideoCompare = _ToolPath($g_sPathVideoCompare, "VideoCompare")
	$g_sPathSync = _ToolPath($g_sPathSync, "Sync")
	$g_sPathFFmpeg = _ToolPath($g_sPathFFmpeg, "FFmpeg")
	; Лончер может лежать не в папке video-compare: ffmpeg найдёт свои av*-DLL
	; только рядом с video-compare.exe
	If Not FileExists($g_sPathFFmpeg) Then
		Local $sVcFFmpeg = _PathGetDir($g_sPathVideoCompare) & "\ffmpeg.exe"
		If FileExists($sVcFFmpeg) Then $g_sPathFFmpeg = $sVcFFmpeg
	EndIf
EndFunc   ;==>_ResolveToolPaths


; Путь к инструменту: по умолчанию, а если там его нет – из секции Tools в ini
Func _ToolPath($sDefault, $sIniKey)
	If FileExists($sDefault) Then Return $sDefault
	Local $sIniPath = _NormalizePath(IniRead($gc_sPathIni, "Tools", $sIniKey, ""))
	Return FileExists($sIniPath) ? $sIniPath : $sDefault
EndFunc   ;==>_ToolPath


; Версия video-compare для шапки. Определяется через --version, только если в ini пусто
Func _InitVcVersion()
	$g_sVcVersion = IniRead($gc_sPathIni, "Settings", "VcVersion", "")
	If $g_sVcVersion <> "" Then Return
	$g_sVcVersion = _DetectVcVersion()
	If $g_sVcVersion <> "" Then IniWrite($gc_sPathIni, "Settings", "VcVersion", $g_sVcVersion)
EndFunc   ;==>_InitVcVersion


; Версия из "video-compare 20260502-valencia" в виде '20260502 "valencia"'.
; "" – не запустилось или вывод не распознан.
Func _DetectVcVersion()
	Local $iPid = Run('"' & $g_sPathVideoCompare & '" --version', @ScriptDir, @SW_HIDE, $STDOUT_CHILD)
	If @error Or Not $iPid Then Return ""
	ProcessWaitClose($iPid, 5)
	Local $aMatch = StringRegExp(StdoutRead($iPid), '(?im)video-compare\s+(\S+)', $STR_REGEXPARRAYMATCH)
	If @error Then Return ""
	Local $sRaw = $aMatch[0]
	; Дата и кодовое имя разделены первым дефисом
	Local $iDash = StringInStr($sRaw, "-")
	If $iDash > 0 Then Return StringLeft($sRaw, $iDash - 1) & ' "' & StringTrimLeft($sRaw, $iDash) & '"'
	Return $sRaw
EndFunc   ;==>_DetectVcVersion


; Дата сборки video-compare (ГГГГММДД) из версии '20260828 "reykjavik"'. 0 – не распознана,
; такую версию считаем свежей: вслепую резать ключи хуже, чем передать лишний.
Func _VcBuildDate()
	Local $aMatch = StringRegExp($g_sVcVersion, '(\d{8})', $STR_REGEXPARRAYMATCH)
	If @error Then Return 0
	Return Int($aMatch[0])
EndFunc   ;==>_VcBuildDate


; ============================================================
; Локализация
; ============================================================

Func Lang($sSection, $sKey, $sDefault = "")
	Local $sFullKey = $sSection & "." & $sKey
	If MapExists($g_mLang, $sFullKey) Then Return $g_mLang[$sFullKey]
	Return $sDefault
EndFunc   ;==>Lang


Func _InitLanguage()
	$g_sCurrentLang = IniRead($gc_sPathIni, "Settings", "Language", "")
	; Пусто – по языку системы: русский и украинский получают русский интерфейс
	If $g_sCurrentLang = "" Then $g_sCurrentLang = (@OSLang = "0419" Or @OSLang = "0422") ? "Russian" : "English"
	$g_sLangFile = @ScriptDir & "\Lang\" & $g_sCurrentLang & ".lng"
	If Not FileExists($g_sLangFile) Then
		$g_sLangFile = @ScriptDir & "\Lang\English.lng"
		$g_sCurrentLang = "English"
	EndIf
	_LoadLangFile()
EndFunc   ;==>_InitLanguage


Func _LoadLangFile()
	Local $mEmpty[]
	$g_mLang = $mEmpty
	If $g_sLangFile = "" Or Not FileExists($g_sLangFile) Then Return
	Local $hFile = FileOpen($g_sLangFile, $FO_UTF8_NOBOM)
	If $hFile = -1 Then Return
	Local $sContent = FileRead($hFile)
	FileClose($hFile)
	Local $aLines = StringSplit(StringStripCR($sContent), @LF)
	Local $sSection = ""
	For $i = 1 To $aLines[0]
		Local $sLine = StringStripWS($aLines[$i], $STR_STRIPLEADING + $STR_STRIPTRAILING)
		If $sLine = "" Or StringLeft($sLine, 1) = ";" Then ContinueLoop
		If StringLeft($sLine, 1) = "[" And StringRight($sLine, 1) = "]" Then
			$sSection = StringMid($sLine, 2, StringLen($sLine) - 2)
			ContinueLoop
		EndIf
		Local $iEq = StringInStr($sLine, "=")
		If $iEq = 0 Or $sSection = "" Then ContinueLoop
		Local $sKey = StringStripWS(StringLeft($sLine, $iEq - 1), $STR_STRIPTRAILING)
		$g_mLang[$sSection & "." & $sKey] = StringMid($sLine, $iEq + 1)
	Next
EndFunc   ;==>_LoadLangFile


Func _ApplyLanguage()
	GUICtrlSetData($g_iLabel1, Lang("GUI", "File1", "File 1"))
	GUICtrlSetData($g_iLabel2, Lang("GUI", "File2", "File 2"))
	_SetTip($g_iButtonChoose1, Lang("GUI", "Choose", "Choose"))
	_SetTip($g_iButtonChoose2, Lang("GUI", "Choose", "Choose"))
	_SetTip($g_iButtonSwap, Lang("GUI", "SwapTip", "Swap files"))
	_SetTip($g_iButtonSettings, Lang("GUI", "Settings", "Settings"))
	_SetTip($g_iButtonKeys, Lang("Hotkeys", "1", "Video-compare hotkeys"))
	GUICtrlSetData($g_iLabelCompare, Lang("GUI", "CompareMode", "Mode"))
	_SkinBtnSetText($g_iButtonCompare, Lang("GUI", "Compare", "Compare"))
	GUICtrlSetData($g_iLabelOffset, Lang("GUI", "Offset", "Offset"))
	GUICtrlSetData($g_iLabelOffsetUnits, Lang("GUI", "OffsetUnits", "ms"))
	GUICtrlSetData($g_iLabelCommand, Lang("GUI", "TabCommand", "Command"))
	GUICtrlSetData($g_iLabelCommandHint, Lang("GUI", "CommandHint", "Generated automatically. Editable."))

	; Ширина сегментов зависит от подписей, поэтому раскладка считается заново
	_SegSetTexts($g_iSegCompare, Lang("GUI", "SegDirect", "Direct") & "|" & Lang("GUI", "SegVertical", "Vertical"))
	_SegSetTexts($g_iSegFit, Lang("GUI", "SegNative", "Native") & "|" & Lang("GUI", "SegCrop", "Cropped"))
	_SegSetTexts($g_iSegOffset, Lang("GUI", "OffsetAuto", "Auto") & "|" & Lang("GUI", "OffsetManual", "Manual"))
	_UpdateModeDesc()
	_UpdateModeTips()
	_LayoutBody()

	Local $aRows = _HotkeyRows()
	_SkinHeaderSetText($g_iHeaderPic, $gc_sAppName, _HeaderSubtitle())
	_SkinHotkeysSetRows($g_iHotkeysPanel, Lang("Hotkeys", "1", "Video-compare hotkeys"), $aRows)
	_UpdateFilesInfo()
	_ApplyTheme() ; заодно переводит статус поиска
	_ApplyUpdateNotice() ; пометка об обновлении в заголовке
EndFunc   ;==>_ApplyLanguage


; ============================================================
; Тема
; ============================================================

Func _InitTheme()
	; Рамки полей и кнопок рисует скин: вторая рамка от UDF дала бы двойной контур
	_GUIDarkTheme_CtrlBorderSet(False, False)

	; Пусто (первый запуск) или «System» из прежних версий – тема системы, запоминаем в ini
	$g_sTheme = IniRead($gc_sPathIni, "Settings", "Theme", "")
	If $g_sTheme <> "Light" And $g_sTheme <> "Dark" Then
		$g_sTheme = _ReadSystemDarkMode() ? "Dark" : "Light"
		IniWrite($gc_sPathIni, "Settings", "Theme", $g_sTheme)
	EndIf
	_ResolveDarkMode()
EndFunc   ;==>_InitTheme


; Возвращает выставленный режим: вызывающему он нужен сразу (_MainGUI).
Func _ResolveDarkMode()
	$g_bDarkMode = ($g_sTheme = "Dark")
	Return $g_bDarkMode
EndFunc   ;==>_ResolveDarkMode


Func _ReadSystemDarkMode()
	Local $vVal = RegRead("HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Themes\Personalize", _
			"AppsUseLightTheme")
	If @error Then Return False
	Return ($vVal = 0)
EndFunc   ;==>_ReadSystemDarkMode


; Палитра одна на окно и живёт в SkinCore: подписи и разделители берут цвета оттуда же,
; что и скин, иначе тона разъезжаются.
Func _SetPalette()
	_SkinSetTheme($g_bDarkMode)
	$g_iClrBg = $g_iSkinBg
	$g_iClrFg = $g_iSkinText2   ; подписи строк
	$g_iClrInfo = $g_iSkinText3 ; пояснения под строками
	$g_iClrSep = $g_iSkinDivider
EndFunc   ;==>_SetPalette


Func _ApplyTheme()
	_ResolveDarkMode()
	_SetPalette()

	; Сравниваем с темой, реально применённой к GUI: $g_bDarkMode мог обновить вызывающий
	If Not $g_bThemeInitialized Or $g_bAppliedDark <> $g_bDarkMode Then
		; _GUIDarkTheme_SwitchTheme не годится: направление он берёт из системной темы.
		; Его очистку повторяем сами: кисти и перья UDF кеширует, и без неё они
		; остались бы старого цвета.
		If $g_bThemeInitialized Then
			__GUIDarkTheme_SubclassCleanup()
			__GUIDarkTheme_BrushCleanup()
			__GUIDarkTheme_PenCleanup()
		EndIf
		If $g_bDarkMode Then
			; Фон Edit у UDF свой (0x383838) и пятном проступал внутри рамки скина
			$__DM_g_iCtrlBkColorDark = $g_iSkinCtrlBg
			_GUIDarkTheme_ApplyDark($g_hGui)
		Else
			_GUIDarkTheme_ApplyLight($g_hGui)
		EndIf
		$g_bThemeInitialized = True
		$g_bAppliedDark = $g_bDarkMode
		; UDF переустанавливает свой WM_CTLCOLOREDIT, без делегата пропадёт drag-подсветка
		GUIRegisterMsg($WM_CTLCOLOREDIT, "_OnEvent_WM_CTLCOLOREDIT")
	EndIf

	; Фон окна в обеих темах: иначе под непрозрачными подписями проступает другой тон
	GUISetBkColor($g_iClrBg, $g_hGui)

	; UDF делает static прозрачными, и текст при перерисовке накладывался бы сам на себя.
	; Непрозрачный фон цвета окна выглядит так же, но стирается сплошной кистью.
	Local $aBodyLabels[5] = [$g_iLabel1, $g_iLabel2, $g_iLabelCompare, $g_iLabelOffset, $g_iLabelCommand]
	For $iLbl In $aBodyLabels
		GUICtrlSetBkColor($iLbl, $g_iClrBg)
		GUICtrlSetColor($iLbl, $g_iClrFg)
	Next
	Local $aSubLabels[7] = [$g_iLabelInfo1, $g_iLabelInfo2, $g_iLabelModeDesc, $g_iLabelFitDesc, _
			$g_iLabelCommandHint, $g_iLabelSyncStatus, $g_iLabelOffsetUnits]
	For $iLbl In $aSubLabels
		GUICtrlSetBkColor($iLbl, $g_iClrBg)
		GUICtrlSetColor($iLbl, $g_iClrInfo)
	Next
	Local $aSeparators[3] = [$g_iSeparatorTop, $g_iSepFiles, $g_iSepCompare]
	For $iSep In $aSeparators
		GUICtrlSetBkColor($iSep, $g_iClrSep)
	Next

	; Цвет статуса зависит и от исхода поиска, и от темы
	If $g_sLastSyncStatus <> "" Then _SetSyncStatus($g_sLastSyncStatus, $g_iLastSyncOffset)

	; Скин нарисован картинками старой палитры
	_UpdateCompareButtonIcon()
	_SkinHeaderRenderAll()
	_SkinHotkeysRenderAll()
	_SkinBtnRenderAll()
	_SkinInputRenderAll()
	_SegRenderAll()

	_RedrawGui()
EndFunc   ;==>_ApplyTheme


; Полная синхронная перерисовка главного окна. RDW_ERASE обязателен:
; без стирания фона текст подписей накладывается на прежний.
Func _RedrawGui()
	_WinAPI_RedrawWindow($g_hGui, 0, 0, BitOR($RDW_INVALIDATE, $RDW_ERASE, $RDW_UPDATENOW, $RDW_ALLCHILDREN))
EndFunc   ;==>_RedrawGui


; ============================================================
; Окно настроек
; ============================================================

; Метрика главного окна, выбор сегментами: штатные combo и radio под тему не красятся.
; Ширина окна по самой широкой строке: сегменты меряются по подписям, а те в языках разные.
Func _SettingsWindow()
	; Длинная подпись обновлений в две строки, иначе колонка подписей отодвинет селекторы
	Local $sUpdatesLabel = _SettingsTwoLines(Lang("Updates", "CheckLabel", "Check for updates"))
	Local $iLabelW = _SettingsLabelW($sUpdatesLabel)
	Local $iCtrlX = $gc_iPadX + $iLabelW + $gc_iGap

	; По центру главного окна. Размер черновой: окончательный ставит _FitGuiToClient
	Local $aMain = WinGetPos($g_hGui)
	Local $iSetX = -1, $iSetY = -1
	If Not @error Then
		$iSetX = $aMain[0] + Int($aMain[2] / 2) - 200
		$iSetY = $aMain[1] + Int($aMain[3] / 2) - 100
	EndIf
	$g_hSettingsGui = GUICreate(Lang("GUI", "Settings", "Settings"), 400, 200, $iSetX, $iSetY, _
			$WS_CAPTION + $WS_SYSMENU, $WS_EX_DLGMODALFRAME, $g_hGui)
	GUISetFont($g_iAppFontSize, 400, 0, $g_sAppFont, $g_hSettingsGui)

	Local $iY = $gc_iPadX
	; Между строками воздуха больше, чем в главном окне: с зазором 12 окно выглядело втиснутым
	Local Const $iRowGap = 20

	; --- Язык. Подписи сегментов – самоназвания языков, они не переводятся ---
	$g_iSettingsLabelLang = _SettingsLabel(Lang("GUI", "Language", "Language"), $iY, $iLabelW)
	$g_iSettingsSegLang = _SegCreate("English|Русский", "", $iCtrlX, $iY, $gc_iRowH, _
			($g_sCurrentLang = "Russian") ? 1 : 0)

	$iY += $gc_iRowH + $iRowGap

	; --- Тема ---
	$g_iSettingsLabelTheme = _SettingsLabel(Lang("GUI", "Theme", "Theme"), $iY, $iLabelW)
	$g_iSettingsSegTheme = _SegCreate( _
			Lang("GUI", "ThemeLight", "Light") & "|" & Lang("GUI", "ThemeDark", "Dark"), _
			"SegLight|SegDark", $iCtrlX, $iY, $gc_iRowH, ($g_sTheme = "Dark") ? 1 : 0)

	$iY += $gc_iRowH + $iRowGap

	; --- Проверка обновлений: да/нет сегментами, справа ручная проверка ---
	$g_iSettingsLabelUpdates = _SettingsLabel($sUpdatesLabel, $iY, $iLabelW, 2)
	$g_iSettingsSegUpdates = _SegCreate( _
			Lang("Updates", "CheckYes", "Yes") & "|" & Lang("Updates", "CheckNo", "No"), _
			"", $iCtrlX, $iY, $gc_iRowH, _UpdateCheckEnabled() ? 0 : 1)

	; Ширина строк – по самому широкому селектору, но не уже строки действий.
	; Все строки тянутся до неё, правый край селекторов встаёт по краю «ОК».
	Local $iCacheW = _SettingsBtnW(Lang("GUI", "ClearCache", "Clear cache"), False)
	Local $iOkW = _SettingsBtnW(Lang("GUI", "OK", "OK"), True)
	Local $iRowW = _Max(_Max(_SegWidth($g_iSettingsSegLang), _SegWidth($g_iSettingsSegTheme)), _
			_SegWidth($g_iSettingsSegUpdates) + $gc_iGap + $gc_iIconBtn)
	Local $iClientW = _Max($iCtrlX + $iRowW + $gc_iPadX, _
			$gc_iPadX * 2 + $iCacheW + $gc_iGap + $iOkW)
	$iRowW = $iClientW - $gc_iPadX - $iCtrlX

	_SegSetWidth($g_iSettingsSegLang, $iRowW)
	_SegSetWidth($g_iSettingsSegTheme, $iRowW)

	; Кнопка проверки у правого края, да/нет занимают остаток
	_SegSetWidth($g_iSettingsSegUpdates, $iRowW - $gc_iGap - $gc_iIconBtn)
	$g_iSettingsButtonCheckNow = _SkinBtnCreate("", "Refresh", 18, _
			$iCtrlX + $iRowW - $gc_iIconBtn, $iY, $gc_iIconBtn, $gc_iIconBtn)
	_SetTip($g_iSettingsButtonCheckNow, Lang("Updates", "CheckNow", "Check now"))
	GUICtrlSetOnEvent($g_iSettingsButtonCheckNow, "_OnEvent_SettingsCheckNow")
	If $g_hUpdateDownload Then _SkinBtnSetEnabled($g_iSettingsButtonCheckNow, False) ; проверка уже идёт

	$g_iSettingsLabelUpdateInfo = GUICtrlCreateLabel("", $iCtrlX + 1, $iY + $gc_iSubTop, _
			$iClientW - $gc_iPadX - $iCtrlX, $gc_iSubH, $SS_NOTIFY)
	_SkinDockFixed($g_iSettingsLabelUpdateInfo)
	GUICtrlSetOnEvent($g_iSettingsLabelUpdateInfo, "_OnEvent_SettingsUpdateLink")
	If $g_hLinkSubclassCB = 0 Then $g_hLinkSubclassCB = DllCallbackRegister("_LinkSubclassProc", _
			"lresult", "hwnd;uint;wparam;lparam;uint_ptr;dword_ptr")
	_WinAPI_SetWindowSubclass(GUICtrlGetHandle($g_iSettingsLabelUpdateInfo), _
			DllCallbackGetPtr($g_hLinkSubclassCB), 1, 0)

	; --- Полоса и строка действий: сброс кеша и «ОК» – действия, а не настройки ---
	Local $iSepY = $iY + $gc_iSubTop + $gc_iSubH + $gc_iBlockGap
	$g_iSettingsSep = GUICtrlCreateLabel("", $gc_iPadX, $iSepY, $iClientW - $gc_iPadX * 2, 1)
	_SkinDockFixed($g_iSettingsSep)

	$iY = $iSepY + $gc_iBlockGap + 1

	; Описание кеша – в подсказке кнопки
	$g_iSettingsButtonClearCache = _SkinBtnCreate(Lang("GUI", "ClearCache", "Clear cache"), "", 0, _
			$gc_iPadX, $iY, $iCacheW, $gc_iRowH, $SKINBTN_TEXT)
	_SetTip($g_iSettingsButtonClearCache, Lang("Cache", "Desc", _
			"Stores video resolutions and detected sync offsets. Clear it if cached data looks stale."))
	GUICtrlSetOnEvent($g_iSettingsButtonClearCache, "_OnEvent_SettingsClearCache")

	$g_iSettingsButtonOk = _SkinBtnCreate(Lang("GUI", "OK", "OK"), "", 0, _
			$iClientW - $gc_iPadX - $iOkW, $iY, $iOkW, $gc_iRowH, $SKINBTN_ACCENT)
	GUICtrlSetOnEvent($g_iSettingsButtonOk, "_OnEvent_SettingsOk")

	GUISetOnEvent($GUI_EVENT_CLOSE, "_OnEvent_SettingsClose", $g_hSettingsGui)

	_FitGuiToClient($g_hSettingsGui, $iClientW, $iY + $gc_iRowH + $gc_iPadX)

	_ApplyThemeToSettingsGui()
	_SettingsUpdateInfoText()

	GUISetState(@SW_SHOW, $g_hSettingsGui)
EndFunc   ;==>_SettingsWindow


; Подпись строки настроек в колонке главного окна. Многострочная встаёт по центру
; блока из строки и пояснения под ней. Докинг прибит: _FitGuiToClient подгоняет
; окно после раскладки и иначе растянул бы подписи вслед за собой.
Func _SettingsLabel($sText, $iY, $iLabelW, $iLines = 1)
	Local $iTop = $iY + 6, $iH = 16
	If $iLines > 1 Then
		$iH = _SkinGdiLineHeight() * $iLines
		$iTop = $iY + Int(($gc_iSubTop + $gc_iSubH - $iH) / 2)
	EndIf

	Local $iCtrl = GUICtrlCreateLabel($sText, $gc_iPadX, $iTop, $iLabelW, $iH)
	_SkinDockFixed($iCtrl)
	Return $iCtrl
EndFunc   ;==>_SettingsLabel


; Две строки с разрывом по пробелу, при котором длинная из них короче всего.
; Подпись из одного слова возвращается как есть.
Func _SettingsTwoLines($sText)
	Local $hFont = _SkinFont()
	Local $sBest = $sText, $iBestW = -1

	Local $iPos = StringInStr($sText, " ")
	While $iPos
		Local $sTop = StringLeft($sText, $iPos - 1), $sBottom = StringMid($sText, $iPos + 1)
		Local $iW = _Max(_SkinTextW($sTop, $hFont), _SkinTextW($sBottom, $hFont))
		If $iBestW < 0 Or $iW < $iBestW Then
			$iBestW = $iW
			$sBest = $sTop & @CRLF & $sBottom
		EndIf
		$iPos = StringInStr($sText, " ", 0, 1, $iPos + 1)
	WEnd
	Return $sBest
EndFunc   ;==>_SettingsTwoLines


; Колонка подписей: как в главном окне, но не уже самой длинной строки.
; $sUpdates – подпись обновлений, уже разбитая на строки.
Func _SettingsLabelW($sUpdates)
	Local $hFont = _SkinFont()
	Local $sTexts = Lang("GUI", "Language", "Language") & @CRLF & Lang("GUI", "Theme", "Theme") & @CRLF & $sUpdates
	Local $aLines = StringSplit($sTexts, @CRLF, $STR_ENTIRESPLIT + $STR_NOCOUNT)

	Local $iW = $gc_iLabelW
	For $sLine In $aLines
		Local $iTextW = _SkinTextW($sLine, $hFont) + 4
		If $iTextW > $iW Then $iW = $iTextW
	Next
	Return $iW
EndFunc   ;==>_SettingsLabelW


; Ширина кнопки действия: подпись с полями, но не уже 90 –
; короткое «ОК» иначе выглядит обрезком рядом со «Сбросить кеш».
Func _SettingsBtnW($sText, $bAccent)
	Return _Max(90, _SkinTextW($sText, _SkinFont(0, $bAccent)) + $gc_iPadX * 2)
EndFunc   ;==>_SettingsBtnW


; Автопроверка обновлений выключена только значением UpdateCheck=0
Func _UpdateCheckEnabled()
	Return IniRead($gc_sPathIni, "Settings", "UpdateCheck", "1") <> "0"
EndFunc   ;==>_UpdateCheckEnabled


; Всплывающая подсказка. Windows тянет её одной строкой во всю ширину экрана,
; поэтому длинный текст сначала ломается на строки близкой длины.
Func _SetTip($iCtrlID, $sText)
	GUICtrlSetTip($iCtrlID, _WrapText($sText, $gc_iTipWidth))
	; GUICtrlSetTip пересоздаёт окно подсказки, а UDF затемнила только прежние
	If $g_bThemeInitialized And $g_bAppliedDark Then _DarkenTooltips()
EndFunc   ;==>_SetTip


; Тёмная тема всем подсказкам процесса: у AutoIt своё окно подсказки на каждый контрол.
Func _DarkenTooltips()
	Local $aWnd = _WinAPI_EnumProcessWindows(0, False)
	If @error Then Return
	For $i = 1 To $aWnd[0][0]
		If $aWnd[$i][1] = "tooltips_class32" Then _GUIDarkTheme_GUICtrlSetDarkTheme($g_hGui, $aWnd[$i][0], True)
	Next
EndFunc   ;==>_DarkenTooltips


; Ломает текст по пробелам на строки не длиннее $iMax. Строк берём столько, сколько
; требует длина, а ширину подбираем минимальную, при которой их не становится больше:
; иначе жадная укладка оставляет последнюю строку огрызком в пару слов.
Func _WrapText($sText, $iMax = $gc_iTipWidth)
	$sText = StringStripWS($sText, $STR_STRIPLEADING + $STR_STRIPTRAILING + $STR_STRIPSPACES)
	Local $iLen = StringLen($sText)
	If $iLen <= $iMax Then Return $sText

	Local $iLines = Ceiling($iLen / $iMax)
	Local $sWrapped = ""
	For $iWidth = Ceiling($iLen / $iLines) To $iMax
		$sWrapped = __WrapGreedy($sText, $iWidth)
		StringReplace($sWrapped, @CRLF, @CRLF)
		If @extended + 1 <= $iLines Then Return $sWrapped
	Next

	Return __WrapGreedy($sText, $iMax)
EndFunc   ;==>_WrapText


; Жадная укладка слов в строки шириной $iWidth. Слово длиннее строки не режется.
Func __WrapGreedy($sText, $iWidth)
	Local $aWords = StringSplit($sText, " ", $STR_NOCOUNT)
	Local $sOut = "", $sLine = ""

	For $sWord In $aWords
		If $sLine = "" Then
			$sLine = $sWord
		ElseIf StringLen($sLine) + 1 + StringLen($sWord) <= $iWidth Then
			$sLine &= " " & $sWord
		Else
			$sOut &= $sLine & @CRLF
			$sLine = $sWord
		EndIf
	Next

	Return $sOut & $sLine
EndFunc   ;==>__WrapGreedy


Func _ApplyThemeToSettingsGui()
	If $g_hSettingsGui = 0 Then Return

	; Реестр subclass-кнопок у UDF один на процесс, и тема каждого окна его обнуляет.
	; Штатные кнопки главного окна (на Win10 до 24H2) читают его без проверки границ,
	; и тема окна настроек роняла бы AutoIt. Реестр сохраняем вокруг покраски.
	Local $aButtonSubSaved = $__DM_g_aButtonSub
	Local $iButtonCountSaved = $__DM_g_iButtonCount

	_GUIDarkTheme_GUISetDarkTheme($g_hSettingsGui, $g_bDarkMode)
	_GUIDarkTheme_GUICtrlAllSetDarkTheme($g_hSettingsGui, $g_bDarkMode)

	$__DM_g_aButtonSub = $aButtonSubSaved
	$__DM_g_iButtonCount = $iButtonCountSaved

	GUISetBkColor($g_iClrBg, $g_hSettingsGui)

	; Подписи строк – тот же слой, что в главном окне: непрозрачный фон цвета окна
	Local $aLabels[3] = [$g_iSettingsLabelLang, $g_iSettingsLabelTheme, $g_iSettingsLabelUpdates]
	For $iLbl In $aLabels
		GUICtrlSetBkColor($iLbl, $g_iClrBg)
		GUICtrlSetColor($iLbl, $g_iClrFg)
	Next

	; Цвет пояснения (ссылка или дата) ставит _SettingsUpdateInfoText
	GUICtrlSetBkColor($g_iSettingsLabelUpdateInfo, $g_iClrBg)

	; Полосу красим после темы UDF: она переводит static в прозрачный фон
	GUICtrlSetBkColor($g_iSettingsSep, $g_iClrSep)
EndFunc   ;==>_ApplyThemeToSettingsGui


Func _OnEvent_SettingsClose()
	If $g_hSettingsGui = 0 Then Return
	If $g_hLinkSubclassCB Then _WinAPI_RemoveWindowSubclass(GUICtrlGetHandle($g_iSettingsLabelUpdateInfo), _
			DllCallbackGetPtr($g_hLinkSubclassCB), 1)

	; Скин держит HBITMAP и строки реестра модулей, о которых GUIDelete не знает,
	; а освободившийся ControlID AutoIt выдаёт заново
	_SegDelete($g_iSettingsSegLang)
	_SegDelete($g_iSettingsSegTheme)
	_SegDelete($g_iSettingsSegUpdates)
	_SkinBtnDelete($g_iSettingsButtonCheckNow)
	_SkinBtnDelete($g_iSettingsButtonClearCache)
	_SkinBtnDelete($g_iSettingsButtonOk)

	GUIDelete($g_hSettingsGui)
	$g_hSettingsGui = 0
	$g_iSettingsButtonClearCache = 0
	$g_iSettingsButtonCheckNow = 0
	$g_iSettingsButtonOk = 0
EndFunc   ;==>_OnEvent_SettingsClose


Func _OnEvent_SettingsOk()
	If $g_hSettingsGui = 0 Then Return

	; Выбор читаем до удаления окна
	Local $sLangCode = (_SegGetSel($g_iSettingsSegLang) = 1) ? "Russian" : "English"
	Local $bLangChanged = ($sLangCode <> $g_sCurrentLang)
	If $bLangChanged Then IniWrite($gc_sPathIni, "Settings", "Language", $sLangCode)

	Local $sTheme = (_SegGetSel($g_iSettingsSegTheme) = 1) ? "Dark" : "Light"
	Local $bThemeChanged = ($sTheme <> $g_sTheme)
	If $bThemeChanged Then
		$g_sTheme = $sTheme
		IniWrite($gc_sPathIni, "Settings", "Theme", $sTheme)
	EndIf

	IniWrite($gc_sPathIni, "Settings", "UpdateCheck", (_SegGetSel($g_iSettingsSegUpdates) = 1) ? 0 : 1)

	_OnEvent_SettingsClose()

	; _ApplyTheme подхватывает новую $g_sTheme и перерисовывает окно целиком
	If $bLangChanged Then
		_InitLanguage()
		_ApplyLanguage() ; внутри _ApplyTheme
	ElseIf $bThemeChanged Then
		_ApplyTheme()
	EndIf
EndFunc   ;==>_OnEvent_SettingsOk


Func _OnEvent_SettingsClearCache()
	Local $mEmpty[]
	$g_oCache = $mEmpty
	FileDelete($gc_sPathCache)
	_EnsureUtf16File($gc_sPathCache)
	_SkinBtnSetEnabled($g_iSettingsButtonClearCache, False)

	; В режиме «Авто» сдвиг брался из кеша
	If _IsAutoOffset() Then
		GUICtrlSetData($g_iInputOffset, "")
		_SetSyncStatus("NOTRUN")
		_UpdateCommandField()
	EndIf
EndFunc   ;==>_OnEvent_SettingsClearCache


; Пояснение под строкой обновлений: ссылка акцентом при найденной версии,
; иначе серая дата последней проверки. Ручная проверка работает и при выключенной авто.
Func _SettingsUpdateInfoText()
	If $g_hSettingsGui = 0 Then Return

	If $g_sAvailableVersion <> "" Then
		GUICtrlSetData($g_iSettingsLabelUpdateInfo, _
				StringFormat(Lang("Updates", "AvailableVersion", "Version %s is available"), $g_sAvailableVersion))
		GUICtrlSetColor($g_iSettingsLabelUpdateInfo, $g_iSkinAccent)
	Else
		Local $sLast = IniRead($gc_sPathIni, "Settings", "LastUpdateCheck", "")
		If $sLast = "" Then $sLast = Lang("Updates", "LastCheckNever", "never")
		GUICtrlSetData($g_iSettingsLabelUpdateInfo, Lang("Updates", "LastCheck", "Last check:") & " " & $sLast)
		GUICtrlSetColor($g_iSettingsLabelUpdateInfo, $g_iClrInfo)
	EndIf
EndFunc   ;==>_SettingsUpdateInfoText


; ============================================================
; Обновления
; ============================================================

; Версии вида "1.07" покомпонентно как числа: -1, 0 или 1. Недостающий компонент – ноль
Func _CompareVersions($s1, $s2)
	Local $a1 = StringSplit($s1, ".", $STR_NOCOUNT)
	Local $a2 = StringSplit($s2, ".", $STR_NOCOUNT)
	Local $iMax = _Max(UBound($a1), UBound($a2))
	For $i = 0 To $iMax - 1
		Local $iN1 = ($i < UBound($a1)) ? Number($a1[$i]) : 0
		Local $iN2 = ($i < UBound($a2)) ? Number($a2[$i]) : 0
		If $iN1 < $iN2 Then Return -1
		If $iN1 > $iN2 Then Return 1
	Next
	Return 0
EndFunc   ;==>_CompareVersions


; Автопроверка раз в неделю, без всплывающих окон. Найдено обновление – мигает шестерёнка.
Func _CheckUpdates()
	If Not _UpdateCheckEnabled() Then Return

	Local $sLast = IniRead($gc_sPathIni, "Settings", "LastUpdateCheck", "")
	; В ini дата через точки, _DateDiff ждёт YYYY/MM/DD
	If $sLast <> "" And _DateDiff("D", StringReplace($sLast, ".", "/"), _NowCalc()) < 7 Then Return

	_StartUpdateCheck(True)
EndFunc   ;==>_CheckUpdates


; Запускает фоновую загрузку Include\AppConstants.au3 из ветки main: синхронный
; InetRead без сети держал окно мёртвым до таймаута. Результат разбирает
; _UpdateCheckTick. $bPulse – мигать шестерёнкой при находке.
; True – проверка идёт, в том числе начатая раньше.
Func _StartUpdateCheck($bPulse)
	If $g_hUpdateDownload Then Return True

	$g_sUpdateTemp = @TempDir & "\VCLauncher_" & @AutoItPID & ".ver"
	$g_hUpdateDownload = InetGet($gc_sVerCheckUrl, $g_sUpdateTemp, $INET_FORCERELOAD, $INET_DOWNLOADBACKGROUND)
	If @error Or Not $g_hUpdateDownload Then
		$g_hUpdateDownload = 0
		_FinishUpdateCheck("")
		Return False
	EndIf

	$g_bUpdatePulse = $bPulse
	$g_hUpdateTimer = TimerInit()
	AdlibRegister("_UpdateCheckTick", 250)
	Return True
EndFunc   ;==>_StartUpdateCheck


; Ждёт конца загрузки версии. Зависшую загрузку обрывает через $gc_iUpdateTimeoutMs.
Func _UpdateCheckTick()
	Local $bTimeout = TimerDiff($g_hUpdateTimer) > $gc_iUpdateTimeoutMs
	If Not InetGetInfo($g_hUpdateDownload, $INET_DOWNLOADCOMPLETE) And Not $bTimeout Then Return

	AdlibUnRegister("_UpdateCheckTick")
	Local $bOk = Not $bTimeout And InetGetInfo($g_hUpdateDownload, $INET_DOWNLOADSUCCESS)
	InetClose($g_hUpdateDownload)
	$g_hUpdateDownload = 0

	Local $sRemote = "", $hFile = $bOk ? FileOpen($g_sUpdateTemp, $FO_BINARY) : -1
	If $hFile <> -1 Then
		Local $sText = BinaryToString(FileRead($hFile), $SB_UTF8)
		FileClose($hFile) ; иначе _FinishUpdateCheck не удалит файл
		Local $aMatch = StringRegExp($sText, '\$gc_sAppVersion\s*=\s*"([^"]+)"', $STR_REGEXPARRAYMATCH)
		If Not @error Then $sRemote = $aMatch[0]
	EndIf
	_FinishUpdateCheck($sRemote)
EndFunc   ;==>_UpdateCheckTick


; Итог проверки: дата пишется и при сбое сети, новее найденная версия – в заголовок,
; окно настроек, если открыто, получает разблокированную кнопку и свежее пояснение.
Func _FinishUpdateCheck($sRemote)
	FileDelete($g_sUpdateTemp)
	IniWrite($gc_sPathIni, "Settings", "LastUpdateCheck", StringReplace(_NowCalcDate(), "/", "."))

	If $sRemote <> "" And _CompareVersions($gc_sAppVersion, $sRemote) < 0 Then
		$g_sAvailableVersion = $sRemote
		_ApplyUpdateNotice()
		; Открытые настройки и так показывают ссылку, мигать незачем
		If $g_bUpdatePulse And $g_hSettingsGui = 0 Then _StartSettingsPulse()
	EndIf

	If $g_hSettingsGui Then
		_SkinBtnSetEnabled($g_iSettingsButtonCheckNow, True)
		_SettingsUpdateInfoText()
	EndIf
EndFunc   ;==>_FinishUpdateCheck


; Пометка о доступном обновлении в заголовке окна
Func _ApplyUpdateNotice()
	If $g_sAvailableVersion = "" Then Return
	WinSetTitle($g_hGui, "", $gc_sAppName & " " & Lang("Updates", "TitleUpdate", "[Update available]"))
EndFunc   ;==>_ApplyUpdateNotice


; Мигание шестерёнки привлекает внимание к найденному обновлению
Func _StartSettingsPulse()
	If $g_bPulseActive Then Return
	$g_bPulseActive = True
	$g_bPulsePhase = True
	AdlibRegister("_PulseSettingsButton", 600)
EndFunc   ;==>_StartSettingsPulse


Func _StopSettingsPulse()
	If Not $g_bPulseActive Then Return
	AdlibUnRegister("_PulseSettingsButton")
	$g_bPulseActive = False
	_SkinBtnSetOn($g_iButtonSettings, False)
EndFunc   ;==>_StopSettingsPulse


Func _PulseSettingsButton()
	_SkinBtnSetOn($g_iButtonSettings, $g_bPulsePhase)
	$g_bPulsePhase = Not $g_bPulsePhase
EndFunc   ;==>_PulseSettingsButton


Func _OnEvent_SettingsUpdateLink()
	If $g_sAvailableVersion <> "" Then ShellExecute($gc_sReleasesUrl)
EndFunc   ;==>_OnEvent_SettingsUpdateLink


; Курсор-рука над пояснением обновлений, только когда в нём ссылка
Func _LinkSubclassProc($hWnd, $iMsg, $wParam, $lParam, $iId, $dwData)
	#forceref $iId, $dwData
	If $iMsg = $WM_SETCURSOR And $g_sAvailableVersion <> "" Then
		_WinAPI_SetCursor(_WinAPI_LoadCursor(0, $IDC_HAND))
		Return 1
	EndIf
	Return _WinAPI_DefSubclassProc($hWnd, $iMsg, $wParam, $lParam)
EndFunc   ;==>_LinkSubclassProc


; Проверка обновлений вне недельного интервала. Кнопку разблокирует _FinishUpdateCheck
Func _OnEvent_SettingsCheckNow()
	If $g_hSettingsGui = 0 Then Return
	If _StartUpdateCheck(False) Then _SkinBtnSetEnabled($g_iSettingsButtonCheckNow, False)
EndFunc   ;==>_OnEvent_SettingsCheckNow


; ============================================================
; Завершение
; ============================================================

; Освобождает GDI-ресурсы и subclass при любом пути выхода
Func _Cleanup()
	If $g_bPulseActive Then AdlibUnRegister("_PulseSettingsButton")
	If $g_hUpdateDownload Then
		AdlibUnRegister("_UpdateCheckTick")
		InetClose($g_hUpdateDownload)
		FileDelete($g_sUpdateTemp)
	EndIf
	; Subclass ссылки снимается до освобождения своего callback
	If $g_hSettingsGui Then _OnEvent_SettingsClose()
	If $g_hBrushDrag Then _WinAPI_DeleteObject($g_hBrushDrag)
	_SegShutdown()
	_SkinBtnShutdown()
	_SkinInputShutdown()
	_SkinHeaderShutdown()
	_SkinHotkeysShutdown()
	_SkinHandCursorRemove($g_iLabelModeDesc) ; subclass снимается до освобождения callback
	_SkinHandCursorRemove($g_iLabelFitDesc)
	_SkinShutdown()
	If $g_hLinkSubclassCB Then
		DllCallbackFree($g_hLinkSubclassCB)
		$g_hLinkSubclassCB = 0
	EndIf
EndFunc   ;==>_Cleanup
