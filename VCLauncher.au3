#pragma compile(Out, VCLauncher.exe)
#pragma compile(Icon, Assets\Icons\Icon.ico)

#NoTrayIcon
#RequireAdmin

#include <GUIConstantsEx.au3>
#include <WindowsConstants.au3>
#include <WindowsNotifsConstants.au3>
#include <FileConstants.au3>
#include <WinAPI.au3>
#include <Math.au3>
#include <ColorConstants.au3>
#include <EditConstants.au3>
#include <ComboConstants.au3>
#include <ButtonConstants.au3>
#include <APIProcConstants.au3>
#include <APISysConstants.au3>
#include <APIGdiConstants.au3>
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
#include "Include\HeaderHelper.au3"

Opt("GUIOnEventMode", 1)

; Константы приложения ($gc_sAppName и сетевые URL - в Include\AppConstants.au3)
Global $g_sVcVersion = "" ; версия video-compare для шапки; из ini или определяется через --version
Global Const $gc_sPathIni = @ScriptDir & '\VCLauncher.ini'
Global Const $gc_sPathCache = @ScriptDir & '\VCLauncher.cache'
Global Const $gc_aSupportedExtensions[] = [ _
        "mp4", "m4v", "mov", "mkv", "webm", "avi", "wmv", "asf", "flv", "f4v", "mpg", "mpeg", "mp2", "m2v", _
        "ts", "m2ts", "mts", "mxf", "vob", "3gp", "3g2", "ogv", "ogg", "dv", "divx", "rm", "rmvb", "gif", "vpy"]
Global Const $gc_iGuiWidth = 548
Global Const $gc_iHeaderH = 120 ; высота шапки с логотипом
; Картинка-шапка ýже окна на ширину кнопки настроек: кнопка ложится на фон GUI,
; а не на Pic (иначе её иконка пропадает после перерисовки шапки).
Global Const $gc_iHeaderW = $gc_iGuiWidth - 44
Global Const $gc_iGuiHeight = $gc_iHeaderH + 289 ; высота клиента (+4px воздуха после полоски)
; Фон окна по темам. Используется и в палитре (_SetPalette), и в фоне шапки (_RenderHeader),
; чтобы полоса справа под кнопкой настроек не отличалась по тону.
Global Const $gc_iClrBgDark = 0x1E1E1E
Global Const $gc_iClrBgLight = 0xF0F0F0
; Акцентный синий (ссылка обновления, пульс шестерёнки) и цвета светлой палитры
Global Const $gc_iClrAccent = 0x0078D4
Global Const $gc_iClrSepLight = 0xC0C0C0, $gc_iClrInfoLight = 0x808080
; Переменные путей к инструментам (загружаются из ini)
Global $g_sPathVideoCompare = @ScriptDir & '\video-compare.exe'
Global $g_sPathSync = @ScriptDir & '\Sync\dist\Sync.exe'
Global $g_sPathFFmpeg = @ScriptDir & '\Sync\dist\FFmpeg.exe'

; Пропуск N секунд от начала видео при поиске сдвига (Sync.exe --skip N)
Global $g_iSyncSkipSec = Int(IniRead($gc_sPathIni, "Settings", "SyncSkipSec", 300))
Global $g_sSyncMethod = IniRead($gc_sPathIni, "Settings", "SyncMethod", "audio") ; audio | video
; Максимальное время работы Sync.exe в секундах — по истечении процесс убивается
Global $g_iSyncTimeoutSec = Int(IniRead($gc_sPathIni, "Settings", "SyncTimeoutSec", 60))
; Подозрительно большой сдвиг (мс): свыше — перепроверяем встречным методом по видео
Global $g_iSyncSuspectMs = Int(IniRead($gc_sPathIni, "Settings", "SyncSuspectMs", 10000))
; Допуск совпадения audio и video при перепроверке (мс)
Global $g_iSyncVerifyTolMs = Int(IniRead($gc_sPathIni, "Settings", "SyncVerifyTolMs", 500))

; Создание ini и cache по умолчанию, если отсутствуют
_EnsureIniDefaults()
_EnsureUtf16File($gc_sPathCache)

; Ссылки на элементы GUI
Global $g_hGui, $g_iInput1, $g_iInput2, $g_iButtonChoose1, $g_iButtonChoose2, $g_iButtonSwap, $g_iButtonCompare, $g_iRadioDirect, $g_iRadioVertical
Global $g_iLogoPic, $g_hLogoBitmap = 0 ; картинка-логотип в шапке и её HBITMAP для освобождения
Global $g_hCompareImgList = 0 ; HIMAGELIST иконки кнопки «Сравнить» для освобождения при замене
Global $g_hSettingsImgList = 0, $g_hSwapImgList = 0 ; HIMAGELIST иконок кнопок настроек/свопа
Global $g_sAppFont = "MS Shell Dlg 2", $g_iAppFontSize = 9
; Кеш в памяти: разрешения видео и сдвиги sync. Ключ включает mtime —
; автоматически инвалидируется, если файл на диске заменён.
Global $g_oCache[]
Global $g_iLabel1, $g_iLabel2, $g_iLabelInfo1, $g_iLabelInfo2, $g_iLabelCompare, $g_iLabelCommand
Global $g_iEditCommand, $g_iButtonSettings ; кнопка настроек в шапке
Global $g_iSeparator, $g_iSeparatorTop
Global $g_hSettingsGui = 0, $g_iSettingsComboLang, $g_iSettingsComboTheme
Global $g_iSettingsLabelLang, $g_iSettingsLabelTheme
Global $g_iSettingsButtonClearCache = 0
Global $g_iSettingsComboUpdates = 0, $g_iSettingsLabelUpdates = 0 ; combo обновлений + подпись
Global $g_iSettingsButtonCheckNow = 0, $g_hSettingsCheckImgList = 0 ; кнопка ручной проверки + её imagelist
Global $g_iSettingsLabelUpdateInfo = 0 ; одно поле: ссылка «Доступна версия X.XX» или дата проверки
Global $g_hLinkSubclassCB = 0 ; callback subclass поля-ссылки (курсор-рука)
Global $g_iLabelOffset, $g_iRadioOffsetAuto, $g_iRadioOffsetManual, $g_iLabelSyncStatus, $g_iInputOffset

Global $g_sAvailableVersion = "" ; версия из проверки, если доступна новее текущей
Global $g_bPulseActive = False, $g_bPulsePhase = False ; мигание кнопки настроек при наличии обновления

; Состояние последнего sync-запуска (для статус-иконки и MsgBox с деталями)
Global $g_sLastSyncStatus = ""    ; "" | WORKING | OK | NOMATCH | TIMEOUT | ERROR
Global $g_iLastSyncOffset = 0
Global $g_sLastSyncCmd = ""
Global $g_sLastSyncOutput = ""    ; stdout + stderr дочернего Sync.exe

; Кадры спиннера статус-иконки. ChrW для гарантированно UTF-16 в исходнике.
Global Const $gc_aSpinnerFrames[4] = [ChrW(0x25D0), ChrW(0x25D3), ChrW(0x25D1), ChrW(0x25D2)]
Global $g_iSpinnerIdx = 0

Global $g_sLangFile = "", $g_sCurrentLang = "", $g_mLang[]

; Тема оформления
Global $g_sTheme = "System" ; System | Light | Dark
Global $g_bDarkMode = False
Global $g_bThemeInitialized = False ; флаг: применяли ли уже UDF-тему
Global $g_bAppliedDark = False ; тема, реально применённая к GUI (для отслеживания смены)
Global $g_iClrBg, $g_iClrFg, $g_iClrInfo, $g_iClrSep
; Подсветка поля при drag&drop. UDF 3.0.0 монополизирует WM_CTLCOLOREDIT, поэтому
; временную голубую заливку даём через свой делегат-хендлер (см. _OnEvent_WM_CTLCOLOREDIT).
Global $g_hDragHiCtrl = 0 ; HWND подсвечиваемого поля (0 = нет)
Global $g_hBrushDrag = 0  ; кисть подсветки (создаётся лениво)
; Палитра для кастомной отрисовки. Раньше эти глобалы объявлял форк GUIDarkTheme;
; ванильная UDF 3.0.0 их не содержит, поэтому держим их в проекте. Дефолт - тёмные
; значения, _ApplyTheme переопределяет их при смене темы.
Global $COLOR_BG_DARK = 0x1E1E1E, $COLOR_TEXT_LIGHT = 0xCCCCCC, $COLOR_CONTROL_BG = 0x2D2D2D
Global $COLOR_BORDER_LIGHT = 0x555555, $COLOR_BORDER = 0x3C3C3C
; Остальные переменные (чтение и нормализация путей; относительные пути считаем от папки скрипта)
Global $g_sVideoFile1 = _NormalizePath(IniRead($gc_sPathIni, "LastDirs", "Video1", ""))
Global $g_sVideoFile2 = _NormalizePath(IniRead($gc_sPathIni, "LastDirs", "Video2", ""))


_ResolveToolPaths()
_InitLanguage()
_InitTheme()

_CheckToolExists($g_sPathVideoCompare, "video-compare.exe")
_CheckToolExists($g_sPathSync, "Sync.exe")

_InitVcVersion()

OnAutoItExitRegister("_Cleanup")

_MainGUI()
_DefineEvents()

; Фоновая проверка обновлений (окно уже показано в _MainGUI)
_CheckUpdates()

While 1
	Sleep(50)
WEnd


Func _MainGUI()
	$g_hGui = GUICreate($gc_sAppName, $gc_iGuiWidth, $gc_iGuiHeight, -1, -1, $WS_SIZEBOX + $WS_SYSMENU + $WS_MINIMIZEBOX, $WS_EX_ACCEPTFILES)

	_ApplyAppFont()

	; Шапка: логотип слева + шпаргалка горячих клавиш справа (рисуется в _RenderHeader).
	; Pic ýже окна — справа остаётся полоса фона GUI под кнопку настроек.
	$g_iLogoPic = GUICtrlCreatePic("", 0, 0, $gc_iHeaderW, $gc_iHeaderH)

	; Кнопка настроек лежит на фоне GUI правее картинки-шапки, не перекрывая Pic.
	$g_iButtonSettings = GUICtrlCreateButton("", $gc_iGuiWidth - 40, 8, 30, 30) ; иконка-шестерёнка ставится ниже (PNG)
	GUICtrlSetTip($g_iButtonSettings, Lang("GUI", "Settings", "Settings"))

	; Разделитель сразу под шапкой (как перед блоком «Смещение»). Цвет ставит _ApplyTheme.
	$g_iSeparatorTop = GUICtrlCreateLabel("", 0, $gc_iHeaderH, $gc_iGuiWidth, 1)

	Local $iY = $gc_iHeaderH + 16 ; отступ от нижнего края шапки (+4px воздуха после полоски)

	; --- Видео 1 ---
	$g_iLabel1 = GUICtrlCreateLabel(Lang("GUI", "File1", "File 1"), 10, $iY + 3, 74, 20)

	$g_iInput1 = GUICtrlCreateInput("", 90, $iY, $gc_iGuiWidth - 180, 21)
	GUICtrlSetState($g_iInput1, $GUI_DROPACCEPTED)
	If $g_sVideoFile1 <> "" Then
		GUICtrlSetData($g_iInput1, _GetFileName($g_sVideoFile1))
		_ResetInputCaret($g_iInput1)
	EndIf

	$g_iButtonChoose1 = GUICtrlCreateButton(Lang("GUI", "Choose", "Choose"), $gc_iGuiWidth - 82, $iY - 1, 70, 23)

	$g_iLabelInfo1 = GUICtrlCreateLabel(Lang("GUI", "FileNotSelected", "File not selected"), 90, $iY + 24, $gc_iGuiWidth - 180, 15)

	; --- Видео 2 ---
	$g_iLabel2 = GUICtrlCreateLabel(Lang("GUI", "File2", "File 2"), 10, $iY + 54, 74, 20)

	$g_iInput2 = GUICtrlCreateInput("", 90, $iY + 52, $gc_iGuiWidth - 180, 21)
	GUICtrlSetState($g_iInput2, $GUI_DROPACCEPTED)
	If $g_sVideoFile2 <> "" Then
		GUICtrlSetData($g_iInput2, _GetFileName($g_sVideoFile2))
		_ResetInputCaret($g_iInput2)
	EndIf

	$g_iButtonChoose2 = GUICtrlCreateButton(Lang("GUI", "Choose", "Choose"), $gc_iGuiWidth - 82, $iY + 51, 70, 23)

	$g_iLabelInfo2 = GUICtrlCreateLabel(Lang("GUI", "FileNotSelected", "File not selected"), 90, $iY + 76, $gc_iGuiWidth - 180, 15)

	; --- Поменять местами ---
	$g_iButtonSwap = GUICtrlCreateButton("", $gc_iGuiWidth - 60, $iY + 24, 26, 26) ; иконка-свап ставится ниже (PNG)
	GUICtrlSetTip($g_iButtonSwap, Lang("GUI", "SwapTip", "Swap files"))

	; --- Разделитель --- (цвет ставит _ApplyTheme)
	$g_iSeparator = GUICtrlCreateLabel("", 0, $iY + 100, $gc_iGuiWidth, 1)

	; --- Режим сравнения ---
	$g_iLabelCompare = GUICtrlCreateLabel(Lang("GUI", "CompareMode", "Compare:"), 10, $iY + 143, 74, 20)
	GUIStartGroup()
	$g_iRadioDirect = GUICtrlCreateRadio(Lang("GUI", "CompareDirect", "Direct"), 90, $iY + 141, 130, 20)
	$g_iRadioVertical = GUICtrlCreateRadio(Lang("GUI", "CompareVertical", "Vertical"), 230, $iY + 141, 160, 20)
	GUICtrlSetState($g_iRadioDirect, $GUI_CHECKED)

	; --- Синхронизация ---
	$g_iLabelOffset = GUICtrlCreateLabel(Lang("GUI", "Offset", "Offset:"), 10, $iY + 115, 74, 20)
	GUIStartGroup()
	$g_iRadioOffsetAuto = GUICtrlCreateRadio(Lang("GUI", "OffsetAuto", "Auto"), 90, $iY + 113, 50, 20)
	GUICtrlSetState($g_iRadioOffsetAuto, $GUI_CHECKED)
	$g_iRadioOffsetManual = GUICtrlCreateRadio(Lang("GUI", "OffsetManual", "Manual"), $gc_iGuiWidth - 228, $iY + 113, 72, 20)
	; Статус sync — цветной текст рядом с «Авто»: смещение при OK, краткое слово иначе.
	; SS_NOTIFY чтобы Label принимал клики (детали ошибки).
	$g_iLabelSyncStatus = GUICtrlCreateLabel("", 140, $iY + 115, $gc_iGuiWidth - 376, 20, $SS_NOTIFY)
	GUICtrlSetFont($g_iLabelSyncStatus, $g_iAppFontSize, 400, 0, $g_sAppFont)
	$g_iInputOffset = GUICtrlCreateInput("", $gc_iGuiWidth - 150, $iY + 113, 60, 20)
	GUICtrlSetState($g_iInputOffset, $GUI_DISABLE)

	; --- Команда ---
	$g_iLabelCommand = GUICtrlCreateLabel(Lang("GUI", "TabCommand", "Command"), 10, $iY + 172, 70, 20)
	$g_iEditCommand = GUICtrlCreateEdit("", 90, $iY + 171, $gc_iGuiWidth - 180, 70, BitOR($ES_MULTILINE, $ES_AUTOVSCROLL, $WS_VSCROLL))

	; --- Кнопка «Сравнить» ---
	$g_iButtonCompare = GUICtrlCreateButton(Lang("GUI", "Compare", "Compare"), $gc_iGuiWidth - 82, $iY + 171, 70, 70, $BS_MULTILINE)
	; иконки кнопок назначаются в _ApplyTheme (_ApplyButtonIcons) — зависят от темы

	_SetCtrlResizing()
	_RenderHeader()
	_ApplyTheme()

	GUISetState(@SW_SHOW)

	; Отложенная загрузка: окно уже видно, теперь подтягиваем данные
	_TryFillCachedOffset()
	_UpdateFilesInfo()
EndFunc   ;==>_MainGUI


Func _DefineEvents()
	GUICtrlSetOnEvent($g_iButtonChoose1, "_OnEvent_ButtonChoose")
	GUICtrlSetOnEvent($g_iButtonChoose2, "_OnEvent_ButtonChoose")
	GUICtrlSetOnEvent($g_iButtonSwap, "_OnEvent_ButtonSwap")
	GUICtrlSetOnEvent($g_iButtonCompare, "_OnEvent_ButtonCompare")
	GUICtrlSetOnEvent($g_iButtonSettings, "_OnEvent_ButtonSettings")
	GUICtrlSetOnEvent($g_iRadioOffsetAuto, "_OnEvent_RadioOffsetMode")
	GUICtrlSetOnEvent($g_iRadioOffsetManual, "_OnEvent_RadioOffsetMode")
	GUICtrlSetOnEvent($g_iLabelSyncStatus, "_OnEvent_LabelSyncStatusClick")
	GUICtrlSetOnEvent($g_iRadioDirect, "_OnEvent_RadioChanged")
	GUICtrlSetOnEvent($g_iRadioVertical, "_OnEvent_RadioChanged")
	GUISetOnEvent($GUI_EVENT_CLOSE, "_OnEvent_GUI_EVENT_CLOSE")
	GUISetOnEvent($GUI_EVENT_DROPPED, "_OnEvent_GUI_EVENT_DROPPED")

	GUIRegisterMsg($WM_GETMINMAXINFO, "_OnEvent_WM_GETMINMAXINFO")
	GUIRegisterMsg($WM_COMMAND, "_OnEvent_WM_COMMAND")
	GUIRegisterMsg($WM_DROPFILES, "_OnEvent_WM_DROPFILES")
EndFunc   ;==>_DefineEvents


Func _OnEvent_WM_GETMINMAXINFO($hWnd, $iMsg, $wParam, $lParam)
	#forceref $iMsg, $wParam
	If $hWnd <> $g_hGui Then Return $GUI_RUNDEFMSG

	Local $tMMI = DllStructCreate( _
			"int reserved1;int reserved2;" & _
			"int MaxSizeX;int MaxSizeY;" & _
			"int MaxPositionX;int MaxPositionY;" & _
			"int MinTrackSizeX;int MinTrackSizeY;" & _
			"int MaxTrackSizeX;int MaxTrackSizeY", $lParam)
	DllStructSetData($tMMI, "MinTrackSizeX", $gc_iGuiWidth + 14)
	DllStructSetData($tMMI, "MinTrackSizeY", $gc_iGuiHeight + 14)

	Return 0
EndFunc   ;==>_OnEvent_WM_GETMINMAXINFO


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

	Local $sFile = FileOpenDialog($sTitle, _PathGetDir($sCurrentFile), _GetVideoExtensionsFilter(), 1, _GetFileName($sCurrentFile))
	If Not @error And FileExists($sFile) Then
		If $iButtonID = $g_iButtonChoose1 Then
			$g_sVideoFile1 = $sFile
		Else
			$g_sVideoFile2 = $sFile
		EndIf
		GUICtrlSetData($iInputCtrl, _GetFileName($sFile))
		IniWrite($gc_sPathIni, "LastDirs", $sIniKey, $sFile)
		GUICtrlSetData($g_iInputOffset, "")
		_TryFillCachedOffset()
		_UpdateFilesInfo()
	EndIf
EndFunc   ;==>_OnEvent_ButtonChoose


Func _OnEvent_ButtonCompare()
	If Not FileExists($g_sVideoFile1) Or Not FileExists($g_sVideoFile2) Then Return

	; Индикация «идёт запуск»: блокируем кнопку и меняем её надпись
	GUICtrlSetState($g_iButtonCompare, $GUI_DISABLE)
	GUICtrlSetData($g_iButtonCompare, Lang("GUI", "Running", "Running"))

	; Синхронизация в автоматическом режиме
	If GUICtrlRead($g_iRadioOffsetAuto) = $GUI_CHECKED Then
		_SetSyncStatus("WORKING")
		Local $aSync = _GetSyncOffset($g_sVideoFile1, $g_sVideoFile2)
		GUICtrlSetData($g_iInputOffset, ($aSync[0] = "OK") ? $aSync[1] : "")
		_SetSyncStatus($aSync[0], $aSync[1])
		_UpdateCommandField()
	EndIf

	Local $sCmdLine = StringReplace(StringReplace(GUICtrlRead($g_iEditCommand), @CR, " "), @LF, " ")
	$sCmdLine = StringStripWS($sCmdLine, 3)
	If $sCmdLine <> "" Then _RunVideoCompare($sCmdLine)

	; Возвращаем кнопку в исходное состояние
	GUICtrlSetData($g_iButtonCompare, Lang("GUI", "Compare", "Compare"))
	_UpdateFilesInfo()
EndFunc   ;==>_OnEvent_ButtonCompare


Func _OnEvent_ButtonSettings()
	If $g_hSettingsGui <> 0 Then Return
	_StopSettingsPulse() ; пользователь обратил внимание — гасим мигание
	_SettingsWindow()
EndFunc   ;==>_OnEvent_ButtonSettings


Func _OnEvent_ButtonSwap()
	Local $sTmp = $g_sVideoFile1
	$g_sVideoFile1 = $g_sVideoFile2
	$g_sVideoFile2 = $sTmp

	IniWrite($gc_sPathIni, "LastDirs", "Video1", $g_sVideoFile1)
	IniWrite($gc_sPathIni, "LastDirs", "Video2", $g_sVideoFile2)

	GUICtrlSetData($g_iInput1, $g_sVideoFile1 = "" ? "" : _GetFileName($g_sVideoFile1))
	_ResetInputCaret($g_iInput1)
	GUICtrlSetData($g_iInput2, $g_sVideoFile2 = "" ? "" : _GetFileName($g_sVideoFile2))
	_ResetInputCaret($g_iInput2)

	; Инвертируем сдвиг вместо повторного поиска
	Local $sOffset = GUICtrlRead($g_iInputOffset)
	If $sOffset <> "" Then
		GUICtrlSetData($g_iInputOffset, -Int($sOffset))
	EndIf
	; Статус-текст «Авто» показывает то же смещение — инвертируем и его
	If $g_sLastSyncStatus = "OK" Then _SetSyncStatus("OK", -Int($g_iLastSyncOffset))
	_UpdateFilesInfo()
EndFunc   ;==>_OnEvent_ButtonSwap


Func _OnEvent_RadioChanged()
	_UpdateCompareButtonIcon()
	_UpdateCommandField()
EndFunc   ;==>_OnEvent_RadioChanged


Func _OnEvent_RadioOffsetMode()
	If GUICtrlRead($g_iRadioOffsetAuto) = $GUI_CHECKED Then
		GUICtrlSetState($g_iInputOffset, $GUI_DISABLE)
		_TryFillCachedOffset()
	Else
		GUICtrlSetState($g_iInputOffset, $GUI_ENABLE)
		_SetSyncStatus("")
	EndIf
	_UpdateCommandField()
EndFunc   ;==>_OnEvent_RadioOffsetMode


Func _OnEvent_GUI_EVENT_CLOSE()
	Exit
EndFunc   ;==>_OnEvent_GUI_EVENT_CLOSE


Func _OnEvent_WM_DROPFILES($hWnd, $iMsg, $wParam, $lParam)
	#forceref $iMsg, $wParam, $lParam
	If $hWnd <> $g_hGui Then Return $GUI_RUNDEFMSG

	; Определяем над каким элементом находится курсор
	Local $aCursorInfo = GUIGetCursorInfo($g_hGui)
	If @error Then Return $GUI_RUNDEFMSG

	Local $iCtrlID = $aCursorInfo[4]

	; Подсвечиваем поле под курсором. Цвет даёт делегат _OnEvent_WM_CTLCOLOREDIT,
	; здесь только помечаем контрол и форсим перерисовку.
	$g_hDragHiCtrl = 0
	If $iCtrlID = $g_iInput1 Then
		$g_hDragHiCtrl = GUICtrlGetHandle($g_iInput1)
	ElseIf $iCtrlID = $g_iInput2 Then
		$g_hDragHiCtrl = GUICtrlGetHandle($g_iInput2)
	EndIf
	If $g_hDragHiCtrl Then _WinAPI_InvalidateRect($g_hDragHiCtrl)

	; Через небольшую задержку восстанавливаем исходный вид после drop события
	AdlibRegister("_RestoreControlsStyle", 100)

	Return $GUI_RUNDEFMSG
EndFunc   ;==>_OnEvent_WM_DROPFILES


; Делегат покраски edit. UDF 3.0.0 перехватывает WM_CTLCOLOREDIT глобально, поэтому
; временную drag-подсветку вставляем здесь, а все прочие поля отдаём в обработчик UDF.
Func _OnEvent_WM_CTLCOLOREDIT($hWnd, $iMsg, $wParam, $lParam)
	If $g_hDragHiCtrl <> 0 And $lParam = $g_hDragHiCtrl Then
		If Not $g_hBrushDrag Then $g_hBrushDrag = _WinAPI_CreateSolidBrush(_WinAPI_SwitchColor($COLOR_SKYBLUE))
		_WinAPI_SetBkColor($wParam, _WinAPI_SwitchColor($COLOR_SKYBLUE))
		_WinAPI_SetTextColor($wParam, _WinAPI_SwitchColor(0x000000))
		Return $g_hBrushDrag
	EndIf
	Return __GUIDarkTheme_WM_CTLCOLOR($hWnd, $iMsg, $wParam, $lParam)
EndFunc   ;==>_OnEvent_WM_CTLCOLOREDIT


Func _OnEvent_GUI_EVENT_DROPPED()
	Local $iDropId = @GUI_DropId
	Local $sDropFile = @GUI_DragFile

	; Проверяем, что файл брошен на один из input полей
	If $iDropId <> $g_iInput1 And $iDropId <> $g_iInput2 Then Return

	; Проверяем существование файла
	If Not FileExists($sDropFile) Then Return

	; Проверяем расширение файла
	If Not _IsValidVideoExtension($sDropFile) Then
		Local $sExt = StringLower(StringRegExpReplace($sDropFile, '^.*\.', ''))
		MsgBox(48, $gc_sAppName, Lang("Errors", "UnsupportedFormat", "Unsupported file format:") & " " & $sExt & @CR & @CR & _
				Lang("Errors", "SupportedFormats", "Supported formats:") & " " & _GetVideoExtensionsFilter())
		Return
	EndIf

	; Обновляем данные
	_SetVideoFile($iDropId, $sDropFile)
EndFunc   ;==>_OnEvent_GUI_EVENT_DROPPED


Func _OnEvent_WM_COMMAND($hWnd, $iMsg, $wParam, $lParam)
	#forceref $hWnd, $iMsg, $lParam
	Local $iID = BitAND($wParam, 0xFFFF)
	Local $iNotifyCode = BitShift($wParam, 16)

	; Обрабатываем только событие EN_KILLFOCUS - когда пользователь завершил редактирование
	If $iNotifyCode <> $EN_KILLFOCUS Then Return $GUI_RUNDEFMSG

	; Обработка поля сдвига — обновляем команду при потере фокуса
	If $iID = $g_iInputOffset Then
		_UpdateCommandField()
		Return $GUI_RUNDEFMSG
	EndIf

	If $iID <> $g_iInput1 And $iID <> $g_iInput2 Then Return $GUI_RUNDEFMSG

	Local $sInputValue = GUICtrlRead($iID)
	If $sInputValue = "" Then Return $GUI_RUNDEFMSG

	Local $sFullPath = ""

	; Проверяем, является ли путь абсолютным
	If StringRegExp($sInputValue, "^(?:[A-Za-z]:\\|\\\\|/)") Then
		; Абсолютный путь - используем как есть
		$sFullPath = $sInputValue
	Else
		; Относительный путь - ищем относительно папки текущего видеофайла
		Local $sContextDir = ""
		If $iID = $g_iInput1 And FileExists($g_sVideoFile1) Then
			$sContextDir = _PathGetDir($g_sVideoFile1)
		ElseIf $iID = $g_iInput2 And FileExists($g_sVideoFile2) Then
			$sContextDir = _PathGetDir($g_sVideoFile2)
		EndIf

		If $sContextDir = "" Then
			; Нет контекстной папки - сбрасываем поле
			If $iID = $g_iInput1 Then
				$g_sVideoFile1 = ""
				GUICtrlSetData($iID, "")
			ElseIf $iID = $g_iInput2 Then
				$g_sVideoFile2 = ""
				GUICtrlSetData($iID, "")
			EndIf
			_UpdateFilesInfo()
			Return $GUI_RUNDEFMSG
		EndIf

		$sFullPath = $sContextDir & "\" & $sInputValue
	EndIf

	; Валидация: существование файла и поддерживаемое расширение
	If Not FileExists($sFullPath) Or Not _IsValidVideoExtension($sFullPath) Then
		_SetVideoFile($iID, "")
		Return $GUI_RUNDEFMSG
	EndIf

	; Обновляем данные
	_SetVideoFile($iID, $sFullPath)
	Return $GUI_RUNDEFMSG
EndFunc   ;==>_OnEvent_WM_COMMAND


Func _UpdateFilesInfo()
	Local $bFile1Exists = FileExists($g_sVideoFile1)
	Local $bFile2Exists = FileExists($g_sVideoFile2)

	; Если оба файла существуют, вычисляем crop
	If $bFile1Exists And $bFile2Exists Then
		Local $aInfo1 = _GetVideoInfo($g_sVideoFile1)
		Local $aInfo2 = _GetVideoInfo($g_sVideoFile2)

		; Если хоть одно разрешение не определилось — показываем ошибку и блокируем
		If $aInfo1[0] <= 0 Or $aInfo2[0] <= 0 Then
			GUICtrlSetData($g_iLabelInfo1, _FormatInfoLabel($aInfo1))
			GUICtrlSetData($g_iLabelInfo2, _FormatInfoLabel($aInfo2))
			GUICtrlSetState($g_iButtonCompare, $GUI_DISABLE)
			GUICtrlSetData($g_iEditCommand, "")
			Return
		EndIf

		Local $aCropArgs = _CalculateCropArgs($aInfo1, $aInfo2)

		; Формируем текст для первого файла
		Local $sText1 = Lang("Info", "Resolution", "Resolution") & " " & $aInfo1[0] & "x" & $aInfo1[1]
		If $aCropArgs[2] <> $aInfo1[1] Then
			Local $iCropDiff1 = $aInfo1[1] - $aCropArgs[2]
			$sText1 &= ", " & Lang("Info", "HeightDiff", "height difference") & " " & $iCropDiff1
		EndIf
		GUICtrlSetData($g_iLabelInfo1, $sText1)

		; Формируем текст для второго файла
		Local $sText2 = Lang("Info", "Resolution", "Resolution") & " " & $aInfo2[0] & "x" & $aInfo2[1]
		If $aCropArgs[3] <> $aInfo2[1] Then
			Local $iCropDiff2 = $aInfo2[1] - $aCropArgs[3]
			$sText2 &= ", " & Lang("Info", "HeightDiff", "height difference") & " " & $iCropDiff2
		EndIf
		GUICtrlSetData($g_iLabelInfo2, $sText2)

		GUICtrlSetState($g_iButtonCompare, $GUI_ENABLE)

		; Обновляем поле команды из уже посчитанных инфо/crop (без повторного _GetVideoInfo)
		GUICtrlSetData($g_iEditCommand, _ComputeCommandFrom($aInfo1, $aInfo2, $aCropArgs))
	Else
		; Показываем только разрешение без crop
		If $bFile1Exists Then
			GUICtrlSetData($g_iLabelInfo1, _FormatInfoLabel(_GetVideoInfo($g_sVideoFile1)))
		Else
			GUICtrlSetData($g_iLabelInfo1, Lang("GUI", "FileNotSelected", "File not selected"))
		EndIf

		If $bFile2Exists Then
			GUICtrlSetData($g_iLabelInfo2, _FormatInfoLabel(_GetVideoInfo($g_sVideoFile2)))
		Else
			GUICtrlSetData($g_iLabelInfo2, Lang("GUI", "FileNotSelected", "File not selected"))
		EndIf

		GUICtrlSetState($g_iButtonCompare, $GUI_DISABLE)
		GUICtrlSetData($g_iEditCommand, "")
	EndIf
EndFunc   ;==>_UpdateFilesInfo


; Текст инфо-лейбла для одного файла: «Resolution WxH» или ошибка, если не определилось
Func _FormatInfoLabel($aInfo)
	If $aInfo[0] <= 0 Then Return Lang("Errors", "ResolutionUnknown", "Resolution not detected")
	Return Lang("Info", "Resolution", "Resolution") & " " & $aInfo[0] & "x" & $aInfo[1]
EndFunc   ;==>_FormatInfoLabel


Func _UpdateCommandField()
	GUICtrlSetData($g_iEditCommand, _ComputeCommand())
EndFunc   ;==>_UpdateCommandField


; Строит командную строку video-compare из текущего состояния GUI.
; Возвращает "" если файлов нет или разрешение не определено
; (без валидного разрешения — деление на ноль в _CalculateCropArgs).
Func _ComputeCommand()
	If Not FileExists($g_sVideoFile1) Or Not FileExists($g_sVideoFile2) Then Return ""
	Local $aVideo1Info = _GetVideoInfo($g_sVideoFile1)
	Local $aVideo2Info = _GetVideoInfo($g_sVideoFile2)
	If $aVideo1Info[0] <= 0 Or $aVideo2Info[0] <= 0 Then Return ""
	Local $aCropArgs = _CalculateCropArgs($aVideo1Info, $aVideo2Info)
	Return _ComputeCommandFrom($aVideo1Info, $aVideo2Info, $aCropArgs)
EndFunc   ;==>_ComputeCommand


; Собирает команду из уже посчитанных инфо/crop — без повторного запроса разрешения.
; Вызывающий гарантирует валидные $aVideo*Info (разрешение > 0).
Func _ComputeCommandFrom($aVideo1Info, $aVideo2Info, $aCropArgs)
	Local $sOffset = GUICtrlRead($g_iInputOffset)
	Local $iOffsetMs = ($sOffset <> "") ? Int($sOffset) : 0
	Local $bIsVertical = (GUICtrlRead($g_iRadioVertical) = $GUI_CHECKED)
	Return _BuildVideoCompareCommand($aVideo1Info, $aVideo2Info, $aCropArgs, $bIsVertical, $iOffsetMs)
EndFunc   ;==>_ComputeCommandFrom


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


Func _CalculateCropArgs($aVideo1Info, $aVideo2Info)
	Local $iVideo1Width = $aVideo1Info[0], $iVideo1Height = $aVideo1Info[1]
	Local $iVideo2Width = $aVideo2Info[0], $iVideo2Height = $aVideo2Info[1]

	Local $aCropResult[4]
	$aCropResult[0] = "" ; crop left
	$aCropResult[1] = "" ; crop right
	$aCropResult[2] = $iVideo1Height ; final height 1
	$aCropResult[3] = $iVideo2Height ; final height 2

	; Защита от деления на ноль: если разрешение не определено — без crop
	If $iVideo1Width <= 0 Or $iVideo2Width <= 0 Then Return $aCropResult

	; Если разрешения одинаковые, crop не нужен
	If $iVideo1Width = $iVideo2Width And $iVideo1Height = $iVideo2Height Then
		Return $aCropResult
	EndIf

	; Масштабируем к минимальной ширине
	Local $iTargetWidth = _Min($iVideo1Width, $iVideo2Width)
	Local $nScale1 = $iTargetWidth / $iVideo1Width
	Local $nScale2 = $iTargetWidth / $iVideo2Width
	Local $iScaledH1 = Round($iVideo1Height * $nScale1)
	Local $iScaledH2 = Round($iVideo2Height * $nScale2)
	Local $iDiffHeight = Abs($iScaledH1 - $iScaledH2)

	If $iDiffHeight > 0 Then
		If $iScaledH1 > $iScaledH2 Then
			; Crop видео 1 снизу
			Local $iOrigDiff = Round($iDiffHeight / $nScale1)
			$aCropResult[0] = "-l crop=iw:ih-" & $iOrigDiff & " "
			$aCropResult[2] = $iVideo1Height - $iOrigDiff
		Else
			; Crop видео 2 снизу
			Local $iOrigDiff = Round($iDiffHeight / $nScale2)
			$aCropResult[1] = "-r crop=iw:ih-" & $iOrigDiff & " "
			$aCropResult[3] = $iVideo2Height - $iOrigDiff
		EndIf
	EndIf

	Return $aCropResult
EndFunc   ;==>_CalculateCropArgs


; Уникальный идентификатор файла (имя + размер, без пути)
Func _BuildFileId($sFile)
	Return _GetFileName($sFile) & "|" & FileGetSize($sFile)
EndFunc   ;==>_BuildFileId


; Возвращает составной ключ пары файлов для кеша в памяти
Func _BuildPairKey($sFile1, $sFile2)
	Return _BuildFileId($sFile1) & "|" & FileGetTime($sFile1, $FT_MODIFIED, 1) & _
			"|" & _BuildFileId($sFile2) & "|" & FileGetTime($sFile2, $FT_MODIFIED, 1)
EndFunc   ;==>_BuildPairKey


; Возвращает индекс файла из секции [Info], проверяя mtime.
; Если записи нет или mtime устарел — возвращает массив ["", ""]
; Иначе — [индекс, разрешение]
Func _GetCacheInfo($sFile)
	Local $aEmpty[2] = ["", ""]
	Local $sDiskKey = _BuildFileId($sFile)
	Local $sCached = IniRead($gc_sPathCache, "Info", $sDiskKey, "")
	If $sCached = "" Then Return $aEmpty

	; Формат значения: индекс|mtime|разрешение (3 поля)
	Local $aParts = StringSplit($sCached, "|")
	If $aParts[0] < 3 Then Return $aEmpty

	Local $sMtime = FileGetTime($sFile, $FT_MODIFIED, 1)
	If $aParts[2] <> $sMtime Then Return $aEmpty

	Local $aResult[2] = [$aParts[1], $aParts[3]]
	Return $aResult
EndFunc   ;==>_GetCacheInfo


; Сохраняет инфо о файле в [Info], возвращает присвоенный индекс
Func _SaveCacheInfo($sFile, $sResolution)
	Local $sDiskKey = _BuildFileId($sFile)
	Local $sMtime = FileGetTime($sFile, $FT_MODIFIED, 1)

	; Если уже есть запись — обновляем с тем же индексом
	Local $sCached = IniRead($gc_sPathCache, "Info", $sDiskKey, "")
	Local $iIdx
	If $sCached <> "" Then
		$iIdx = Int(StringSplit($sCached, "|")[1])
	Else
		; Берём счётчик из [Meta], инкрементируем
		$iIdx = Int(IniRead($gc_sPathCache, "Meta", "NextId", "1"))
		IniWrite($gc_sPathCache, "Meta", "NextId", $iIdx + 1)
	EndIf

	IniWrite($gc_sPathCache, "Info", $sDiskKey, $iIdx & "|" & $sMtime & "|" & $sResolution)
	Return $iIdx
EndFunc   ;==>_SaveCacheInfo


; Проверяет, является ли значение кеша маркером неудачи (строка-статус)
Func _IsSyncMarker($vValue)
	Return ($vValue = "TIMEOUT" Or $vValue = "NOMATCH" Or $vValue = "ERROR")
EndFunc   ;==>_IsSyncMarker


; Ищет сдвиг в кеше (память + диск, прямая + обратная пара).
; Возвращает массив [$bFound, $iOffset, $sStatus]:
;   [True, 3833, "OK"]       — валидный сдвиг
;   [True, 0, "TIMEOUT"]     — предыдущая попытка тайм-аутила
;   [True, 0, "NOMATCH"]     — предыдущая попытка не нашла совпадений
;   [True, 0, "ERROR"]       — прочая неудача
;   [False, 0, "ERROR"]      — в кеше ничего нет
Func _LookupSyncCache($sFile1, $sFile2)
	Local $aResult[3] = [False, 0, "ERROR"]
	Local $sMemKey = _BuildPairKey($sFile1, $sFile2)

	; Кеш в памяти (прямая пара) — значение уже нормализовано при записи
	If MapExists($g_oCache, $sMemKey) Then Return __DecodeSyncVal($g_oCache[$sMemKey], False)

	; Кеш в памяти (обратная пара — инвертируем сдвиг для OK, маркер копируем)
	Local $sMemKeyRev = _BuildPairKey($sFile2, $sFile1)
	If MapExists($g_oCache, $sMemKeyRev) Then
		Local $aDec = __DecodeSyncVal($g_oCache[$sMemKeyRev], True)
		$g_oCache[$sMemKey] = ($aDec[2] = "OK") ? $aDec[1] : $aDec[2]
		Return $aDec
	EndIf

	; Кеш на диске: находим индексы из [Info] и ищем в [Sync]
	Local $aInfo1 = _GetCacheInfo($sFile1)
	Local $aInfo2 = _GetCacheInfo($sFile2)
	If $aInfo1[0] = "" Or $aInfo2[0] = "" Then Return $aResult

	; Прямая пара
	Local $sSyncVal = IniRead($gc_sPathCache, "Sync", $aInfo1[0] & "|" & $aInfo2[0], "")
	If $sSyncVal <> "" Then
		Local $aDec = __DecodeSyncVal($sSyncVal, False)
		$g_oCache[$sMemKey] = ($aDec[2] = "OK") ? $aDec[1] : $aDec[2]
		Return $aDec
	EndIf

	; Обратная пара
	Local $sSyncValRev = IniRead($gc_sPathCache, "Sync", $aInfo2[0] & "|" & $aInfo1[0], "")
	If $sSyncValRev <> "" Then
		Local $aDec = __DecodeSyncVal($sSyncValRev, True)
		$g_oCache[$sMemKey] = ($aDec[2] = "OK") ? $aDec[1] : $aDec[2]
		Return $aDec
	EndIf

	Return $aResult
EndFunc   ;==>_LookupSyncCache


; Декодирует значение кеша sync в [True, $iOffset, $sStatus].
; Маркер неудачи → [True, 0, "TIMEOUT|NOMATCH|ERROR"], число → [True, ±сдвиг, "OK"].
; $bReverse инвертирует знак OK-сдвига (для обратной пары файлов).
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


; Сохраняет результат sync в кеш (память + диск).
; $vValue — либо Int (успешный сдвиг), либо строка-маркер: TIMEOUT/NOMATCH/ERROR.
Func _SaveSyncCache($sFile1, $sFile2, $vValue)
	Local $sMemKey = _BuildPairKey($sFile1, $sFile2)
	Local $sMemKeyRev = _BuildPairKey($sFile2, $sFile1)

	If _IsSyncMarker($vValue) Then
		$g_oCache[$sMemKey] = $vValue
		$g_oCache[$sMemKeyRev] = $vValue
	Else
		$g_oCache[$sMemKey] = $vValue
		$g_oCache[$sMemKeyRev] = -$vValue
	EndIf

	; Получаем/создаём индексы файлов в [Info]
	Local $aInfo1 = _GetCacheInfo($sFile1)
	Local $aInfo2 = _GetCacheInfo($sFile2)
	; Если файла ещё нет в [Info] — определяем разрешение (из кеша или ffmpeg),
	; чтобы не писать запись с пустым полем (иначе ffmpeg гоняется повторно).
	Local $iIdx1 = ($aInfo1[0] <> "") ? Int($aInfo1[0]) : _SaveCacheInfo($sFile1, _GetVideoResolution($sFile1))
	Local $iIdx2 = ($aInfo2[0] <> "") ? Int($aInfo2[0]) : _SaveCacheInfo($sFile2, _GetVideoResolution($sFile2))

	IniWrite($gc_sPathCache, "Sync", $iIdx1 & "|" & $iIdx2, $vValue)
EndFunc   ;==>_SaveSyncCache


; Возвращает результат sync для пары файлов: массив [$sStatus, $iOffset].
; Сначала смотрит в кеш — если там любой статус (OK или маркер неудачи), возвращает его
; без повторного запуска. Только отсутствие записи триггерит реальный sync.
Func _GetSyncOffset($sFile1, $sFile2)
	Local $aCached = _LookupSyncCache($sFile1, $sFile2)
	If $aCached[0] Then
		Local $aHit[2] = [$aCached[2], $aCached[1]]
		Return $aHit
	EndIf

	; Вызов Sync.exe с таймаутом и прогрессом
	Local $aRun = _RunSyncWithTimeout($sFile1, $sFile2, $g_iSyncTimeoutSec)

	; Подозрительно большой сдвиг перепроверяем по видео (для audio-метода)
	If $aRun[0] = "OK" And Abs($aRun[1]) > $g_iSyncSuspectMs And $g_sSyncMethod <> "video" Then
		$aRun = _VerifyOffsetByVideo($sFile1, $sFile2, $aRun)
	EndIf

	; Сохраняем ЛЮБОЙ исход: валидный сдвиг как число, неудачу как маркер
	If $aRun[0] = "OK" Then
		_SaveSyncCache($sFile1, $sFile2, $aRun[1])
	Else
		_SaveSyncCache($sFile1, $sFile2, $aRun[0])
	EndIf

	Return $aRun
EndFunc   ;==>_GetSyncOffset


; Перепроверяет подозрительно большой сдвиг встречным video-методом.
; Совпало в пределах допуска — возвращаем исходный результат, иначе ["NOMATCH", 0].
Func _VerifyOffsetByVideo($sFile1, $sFile2, $aPrimary)
	Local $aNoMatch[2] = ["NOMATCH", 0]
	Local $aVideo = _RunSyncWithTimeout($sFile1, $sFile2, $g_iSyncTimeoutSec, "video")
	If $aVideo[0] <> "OK" Then Return $aNoMatch
	If Abs($aVideo[1] - $aPrimary[1]) > $g_iSyncVerifyTolMs Then Return $aNoMatch
	Return $aPrimary
EndFunc   ;==>_VerifyOffsetByVideo


; Запускает Sync.exe и ждёт результат с таймаутом.
; Возвращает массив [$sStatus, $iOffset]:
;   ["OK", <число>]   — sync вернул валидный сдвиг (exit 0 + число в stdout)
;   ["TIMEOUT", 0]    — истёк $iTimeoutSec, процесс убит
;   ["NOMATCH", 0]    — Sync.exe завершился с ненулевым exit code (совпадений не найдено)
;   ["ERROR", 0]      — прочие сбои (exit 0, но stdout без числа)
Func _RunSyncWithTimeout($sFile1, $sFile2, $iTimeoutSec, $sMethod = "")
	If $sMethod = "" Then $sMethod = $g_sSyncMethod
	Local $sCmdLine = '"' & $g_sPathSync & '" sync --method ' & $sMethod & ' --v1 "' & $sFile1 & '" --v2 "' & $sFile2 & '" --skip ' & $g_iSyncSkipSec & ' --ffmpeg "' & $g_sPathFFmpeg & '"'
	; Сохраняем для click-details в статус-иконке
	$g_sLastSyncCmd = $sCmdLine
	$g_sLastSyncOutput = ""

	Local $iPid = Run($sCmdLine, "", @SW_HIDE, $STDERR_CHILD + $STDOUT_CHILD)
	Local $hProcess = _WinAPI_OpenProcess($STANDARD_RIGHTS_SYNCHRONIZE + $PROCESS_QUERY_INFORMATION, False, $iPid)
	Local $hTimer = TimerInit()
	Local $iTimeoutMs = $iTimeoutSec * 1000
	Local $sStdout = "", $sStderr = ""
	Local $aTimeout[2] = ["TIMEOUT", 0]
	Local $aError[2] = ["ERROR", 0]

	While ProcessExists($iPid)
		Local $iElapsed = TimerDiff($hTimer)

		If $iElapsed >= $iTimeoutMs Then
			ProcessClose($iPid)
			ConsoleWrite("Sync.exe: таймаут (" & $iTimeoutSec & " сек)" & @CRLF)
			If $hProcess Then _WinAPI_CloseHandle($hProcess)
			$g_sLastSyncOutput = _ComposeSyncOutput($sStdout, $sStderr)
			Return $aTimeout
		EndIf

		; Неблокирующее чтение обоих потоков
		$sStdout &= StdoutRead($iPid)
		$sStderr &= StderrRead($iPid)
		Sleep(100)
	WEnd

	; Дочитываем остатки
	$sStdout &= _DrainPipe($iPid, False)
	$sStderr &= _DrainPipe($iPid, True)

	$g_sLastSyncOutput = _ComposeSyncOutput($sStdout, $sStderr)

	; Читаем код выхода процесса (0 — успех, 2 — NOMATCH)
	Local $iExitCode = 0
	If $hProcess Then
		Local $aRet = DllCall("kernel32.dll", "bool", "GetExitCodeProcess", "handle", $hProcess, "dword*", 0)
		If Not @error And IsArray($aRet) Then $iExitCode = $aRet[2]
		_WinAPI_CloseHandle($hProcess)
	EndIf

	If $iExitCode <> 0 Then
		Local $aNoMatch[2] = ["NOMATCH", 0]
		Return $aNoMatch
	EndIf

	; Exit 0 — в stdout должно быть число
	Local $sTrim = StringStripWS($sStdout, 3)
	If Not StringRegExp($sTrim, "^-?\d+$") Then Return $aError

	Local $aOk[2] = ["OK", Int($sTrim)]
	Return $aOk
EndFunc   ;==>_RunSyncWithTimeout


; Склеивает stdout/stderr с заголовками в один текст для click-details.
Func _ComposeSyncOutput($sStdout, $sStderr)
	Local $sOut = ""
	Local $sStdoutTrim = StringStripWS($sStdout, 3)
	Local $sStderrTrim = StringStripWS($sStderr, 3)
	If $sStdoutTrim <> "" Then $sOut &= "[stdout]" & @CRLF & $sStdoutTrim
	If $sStderrTrim <> "" Then
		If $sOut <> "" Then $sOut &= @CRLF & @CRLF
		$sOut &= "[stderr]" & @CRLF & $sStderrTrim
	EndIf
	Return $sOut
EndFunc   ;==>_ComposeSyncOutput


; Неблокирующе дочитывает остаток потока процесса до @error (буфер пуст и процесс закрыт).
; $bStderr = True — stderr, иначе stdout.
Func _DrainPipe($iPid, $bStderr)
	Local $sOut = "", $sLine
	While 1
		$sLine = $bStderr ? StderrRead($iPid) : StdoutRead($iPid)
		If @error Then ExitLoop
		$sOut &= $sLine
	WEnd
	Return $sOut
EndFunc   ;==>_DrainPipe


Func _TryFillCachedOffset()
	If GUICtrlRead($g_iRadioOffsetAuto) <> $GUI_CHECKED Then Return
	If Not FileExists($g_sVideoFile1) Or Not FileExists($g_sVideoFile2) Then
		_SetSyncStatus("NOTRUN")
		Return
	EndIf

	Local $aCached = _LookupSyncCache($g_sVideoFile1, $g_sVideoFile2)
	If Not $aCached[0] Then
		_SetSyncStatus("NOTRUN")
		Return
	EndIf

	If $aCached[2] = "OK" Then
		GUICtrlSetData($g_iInputOffset, $aCached[1])
	Else
		GUICtrlSetData($g_iInputOffset, "")
	EndIf
	_SetSyncStatus($aCached[2], $aCached[1])
EndFunc   ;==>_TryFillCachedOffset


; Единая точка обновления статус-текста sync. $sStatus: "" | NOTRUN | WORKING | OK | NOMATCH | TIMEOUT | ERROR.
; В лейбл пишем полный текст статуса; %d (смещение) подставляется для OK.
Func _SetSyncStatus($sStatus, $iOffset = 0)
	_SyncSpinnerStop()
	$g_sLastSyncStatus = $sStatus
	$g_iLastSyncOffset = $iOffset

	If $sStatus = "" Then
		GUICtrlSetData($g_iLabelSyncStatus, "")
		Return
	EndIf

	Local $iColor = 0, $sTextKey = ""
	Switch $sStatus
		Case "NOTRUN"
			$iColor = $g_iClrInfo ; неактивный цвет: поиск ещё не запускался
			$sTextKey = "StatusNotRun"
		Case "WORKING"
			$iColor = $g_iClrInfo ; неактивный цвет во время поиска
			$sTextKey = "StatusWorking"
			_SyncSpinnerStart() ; текст рисует _SyncSpinnerRender (спиннер + статус)
		Case "OK"
			$iColor = $g_bDarkMode ? 0x66BB6A : 0x2E7D32
			$sTextKey = "StatusOk"
		Case "NOMATCH"
			$iColor = $g_bDarkMode ? 0xFFCA28 : 0xC68400
			$sTextKey = "StatusNoMatch"
		Case "TIMEOUT"
			$iColor = $g_bDarkMode ? 0xFFA726 : 0xC85A00
			$sTextKey = "StatusTimeout"
		Case "ERROR"
			$iColor = $g_bDarkMode ? 0xEF5350 : 0xC0392B
			$sTextKey = "StatusError"
		Case Else
			Return
	EndSwitch

	If $sStatus <> "WORKING" Then
		GUICtrlSetData($g_iLabelSyncStatus, StringReplace(Lang("Sync", $sTextKey, $sStatus), "%d", $iOffset))
	EndIf
	GUICtrlSetColor($g_iLabelSyncStatus, $iColor)
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


; Кадр спиннера плюс полный текст статуса поиска.
Func _SyncSpinnerRender()
	GUICtrlSetData($g_iLabelSyncStatus, $gc_aSpinnerFrames[$g_iSpinnerIdx] & " " & Lang("Sync", "StatusWorking", "searching…"))
EndFunc   ;==>_SyncSpinnerRender


; Клик по статус-иконке для NOMATCH/TIMEOUT/ERROR — MsgBox с командой и stdout/stderr.
Func _OnEvent_LabelSyncStatusClick()
	If Not _IsSyncMarker($g_sLastSyncStatus) Then Return

	Local $sBody = ""
	If $g_sLastSyncCmd <> "" Then $sBody &= "[command]" & @CRLF & $g_sLastSyncCmd & @CRLF & @CRLF
	If $g_sLastSyncOutput <> "" Then $sBody &= $g_sLastSyncOutput
	If $sBody = "" Then $sBody = Lang("Sync", "NoDetails", "No details available.")

	; MsgBox обрезает очень длинные строки — ограничим, остальное в консоль
	Local Const $iMaxLen = 4000
	If StringLen($sBody) > $iMaxLen Then
		ConsoleWrite($sBody & @CRLF)
		$sBody = StringLeft($sBody, $iMaxLen) & @CRLF & @CRLF & "... (truncated, see console)"
	EndIf

	MsgBox(64, $gc_sAppName & " — sync details", $sBody)
EndFunc   ;==>_OnEvent_LabelSyncStatusClick


Func _BuildVideoCompareCommand($aVideo1Info, $aVideo2Info, $aCropArgs, $bIsVertical, $iOffsetMs = 0)
	Local $sCmdLine = '"' & $g_sPathVideoCompare & '"'

	; Проверяем, нужно ли окно на весь экран
	If _ShouldUseFullscreen($aVideo1Info, $aVideo2Info, $aCropArgs, $bIsVertical) Then
		$sCmdLine &= " -W"
	EndIf

	; Режим сравнения
	If $bIsVertical Then
		$sCmdLine &= " -m vstack"
	EndIf

	; Сдвиг по времени (мс → секунды)
	If $iOffsetMs <> 0 Then
		Local $nOffsetSec = $iOffsetMs / 1000
		$sCmdLine &= " -t " & $nOffsetSec
	EndIf

	; Добавляем crop и пути к файлам
	$sCmdLine &= " " & $aCropArgs[0] & $aCropArgs[1] & _
			'"' & $g_sVideoFile1 & '" "' & $g_sVideoFile2 & '"'

	Return $sCmdLine
EndFunc   ;==>_BuildVideoCompareCommand


Func _ShouldUseFullscreen($aVideo1Info, $aVideo2Info, $aCropArgs, $bIsVertical)
	Local $tDesktopRect = _WinAPI_GetWorkArea()
	Local $iDesktopWidth  = DllStructGetData($tDesktopRect, "Right")  - DllStructGetData($tDesktopRect, "Left")
	Local $iDesktopHeight = DllStructGetData($tDesktopRect, "Bottom") - DllStructGetData($tDesktopRect, "Top")

	Local $iMaxWidth = _Max($aVideo1Info[0], $aVideo2Info[0])
	Local $iMaxHeight = _Max($aCropArgs[2], $aCropArgs[3])

	If $bIsVertical Then
		$iMaxHeight *= 2 ; Вертикальное расположение удваивает высоту
	EndIf

	Return ($iMaxWidth > $iDesktopWidth Or $iMaxHeight > $iDesktopHeight)
EndFunc   ;==>_ShouldUseFullscreen


Func _RunVideoCompare($sCmdLine)
	ConsoleWrite($sCmdLine & @CRLF)

	Local $tSA = DllStructCreate("dword nLength;ptr lpSD;bool bInherit")
	DllStructSetData($tSA, "nLength", DllStructGetSize($tSA))
	DllStructSetData($tSA, "bInherit", True)

	Local $aPipe = DllCall("kernel32.dll", "bool", "CreatePipe", _
			"handle*", 0,    _
			"handle*", 0,    _
			"struct*", $tSA, _
			"dword",   0)
	If @error Or Not $aPipe[0] Then Return
	Local $hRead = $aPipe[1], $hWrite = $aPipe[2]

	DllCall("kernel32.dll", "bool", "SetHandleInformation", "handle", $hRead, "dword", $HANDLE_FLAG_INHERIT, "dword", 0)

	Local $tSI = DllStructCreate("dword cb;ptr lpReserved;ptr lpDesktop;ptr lpTitle;dword dwX;dword dwY;dword dwXSize;dword dwYSize;dword dwXCountChars;dword dwYCountChars;dword dwFillAttribute;dword dwFlags;word wShowWindow;word cbReserved2;ptr lpReserved2;handle hStdInput;handle hStdOutput;handle hStdError")
	DllStructSetData($tSI, "cb", DllStructGetSize($tSI))
	DllStructSetData($tSI, "dwFlags", $STARTF_USESTDHANDLES)
	DllStructSetData($tSI, "hStdOutput", $hWrite)
	DllStructSetData($tSI, "hStdError", $hWrite)

	Local $tPI = DllStructCreate("handle hProcess;handle hThread;dword dwPid;dword dwTid")
	Local $tCmd = DllStructCreate("wchar[" & StringLen($sCmdLine) + 1 & "]")
	DllStructSetData($tCmd, 1, $sCmdLine)

	Local $aCP = DllCall("kernel32.dll", "bool", "CreateProcessW", _
			"ptr", 0, "struct*", $tCmd, "ptr", 0, "ptr", 0, "bool", True, _
			"dword", $CREATE_NO_WINDOW, "ptr", 0, "ptr", 0, "struct*", $tSI, "struct*", $tPI)

	DllCall("kernel32.dll", "bool", "CloseHandle", "handle", $hWrite)

	If @error Or Not $aCP[0] Then
		DllCall("kernel32.dll", "bool", "CloseHandle", "handle", $hRead)
		Return
	EndIf

	Local $sOutput = ""
	Local $tBuf = DllStructCreate("byte[4096]")
	Local $tRead = DllStructCreate("dword")
	Local $tAvail = DllStructCreate("dword")
	Local $hProcess = DllStructGetData($tPI, "hProcess")
	; Неблокирующее чтение: в паузах Sleep() в OnEventMode GUI успевает обрабатывать события,
	; поэтому окно не «зависает», пока работает video-compare.
	While 1
		Local $aPeek = DllCall("kernel32.dll", "bool", "PeekNamedPipe", _
				"handle",  $hRead,  _
				"ptr",     0,       _
				"dword",   0,       _
				"ptr",     0,       _
				"struct*", $tAvail, _
				"ptr",     0)
		If @error Or Not $aPeek[0] Then ExitLoop
		Local $iAvail = DllStructGetData($tAvail, 1)
		If $iAvail > 0 Then
			Local $aRF = DllCall("kernel32.dll", "bool", "ReadFile", _
					"handle",  $hRead, _
					"struct*", $tBuf,  _
					"dword",   4096,   _
					"struct*", $tRead, _
					"ptr",     0)
			If @error Or Not $aRF[0] Then ExitLoop
			Local $iN = DllStructGetData($tRead, 1)
			If $iN = 0 Then ExitLoop
			Local $tChunk = DllStructCreate("char[" & $iN + 1 & "]", DllStructGetPtr($tBuf))
			$sOutput &= StringLeft(DllStructGetData($tChunk, 1), $iN)
		Else
			; Нет данных — проверяем, жив ли процесс
			Local $aWait = DllCall("kernel32.dll", "dword", "WaitForSingleObject", "handle", $hProcess, "dword", 0)
			If Not @error And $aWait[0] = 0 Then
				; Процесс завершился: дочитаем остатки и выходим
				Local $aPeek2 = DllCall("kernel32.dll", "bool", "PeekNamedPipe", _
						"handle",  $hRead,  _
						"ptr",     0,       _
						"dword",   0,       _
						"ptr",     0,       _
						"struct*", $tAvail, _
						"ptr",     0)
				If @error Or Not $aPeek2[0] Or DllStructGetData($tAvail, 1) = 0 Then ExitLoop
			Else
				Sleep(30)
			EndIf
		EndIf
	WEnd

	; Код выхода процесса: 0 — успех, иначе — ошибка
	Local $iExitCode = 0
	Local $aExit = DllCall("kernel32.dll", "bool", "GetExitCodeProcess", "handle", $hProcess, "dword*", 0)
	If Not @error And $aExit[0] Then $iExitCode = $aExit[2]

	DllCall("kernel32.dll", "bool", "CloseHandle", "handle", $hRead)
	DllCall("kernel32.dll", "bool", "CloseHandle", "handle", $hProcess)
	DllCall("kernel32.dll", "bool", "CloseHandle", "handle", DllStructGetData($tPI, "hThread"))

	; video-compare пишет в консоль в OEM-кодировке (CP866 на русской Windows) — конвертируем в Unicode.
	; CP_OEMCP = 1 — кодовая страница консоли по умолчанию.
	If $sOutput <> "" Then $sOutput = _WinAPI_MultiByteToWideChar($sOutput, 1, 0, True)

	; Приводим код выхода к знаковому: -1 читается лучше, чем 4294967295
	Local $iExitSigned = $iExitCode
	If $iExitSigned > 0x7FFFFFFF Then $iExitSigned -= 0x100000000

	ConsoleWrite($sOutput & @CRLF & "[exit=" & $iExitSigned & "]" & @CRLF)

	If $iExitCode <> 0 Then _ShowVideoCompareError($sOutput, $iExitSigned)
EndFunc   ;==>_RunVideoCompare


Func _ShowVideoCompareError($sOutput, $iExitCode)
	Local $sMsg = _ExtractErrorLines($sOutput)
	If $sMsg = "" Then $sMsg = StringStripWS($sOutput, 3)
	If $sMsg = "" Then $sMsg = Lang("Errors", "VideoCompareNoOutput", "No error message in output.")

	; Обрезаем слишком длинный вывод, чтобы MsgBox не разрастался
	Local Const $iMaxLen = 1500
	If StringLen($sMsg) > $iMaxLen Then $sMsg = "..." & StringRight($sMsg, $iMaxLen)

	MsgBox(16, $gc_sAppName, _
			Lang("Errors", "VideoCompareFailed", "video-compare exited with an error") & " (exit " & $iExitCode & ")" & @CRLF & @CRLF & $sMsg)
EndFunc   ;==>_ShowVideoCompareError


Func _ExtractErrorLines($sOutput)
	; Выделяем строки с маркерами ошибок video-compare и FFmpeg
	Local $sNorm = StringReplace(StringReplace($sOutput, @CRLF, @LF), @CR, @LF)
	Local $aLines = StringSplit($sNorm, @LF, 1)
	Local $aMarkers[5] = ["Error:", "Exception", "terminate called", "Assertion", "Fatal"]
	Local $sResult = ""
	For $i = 1 To $aLines[0]
		For $j = 0 To UBound($aMarkers) - 1
			If StringInStr($aLines[$i], $aMarkers[$j]) Then
				$sResult &= $aLines[$i] & @CRLF
				ExitLoop
			EndIf
		Next
	Next
	Return StringStripWS($sResult, 3)
EndFunc   ;==>_ExtractErrorLines


Func _GetVideoResolution($sVideoPath)
	If $sVideoPath = "" Or Not FileExists($sVideoPath) Then Return ""

	Local $iSize = FileGetSize($sVideoPath)
	Local $sMtime = FileGetTime($sVideoPath, $FT_MODIFIED, 1)
	Local $sMemKey = _BuildFileId($sVideoPath) & "|" & $sMtime

	; Кеш в памяти
	If MapExists($g_oCache, $sMemKey) Then Return $g_oCache[$sMemKey]

	; Кеш на диске [Info]: ключ name|size, значение idx|mtime|resolution
	Local $aInfo = _GetCacheInfo($sVideoPath)
	If $aInfo[1] <> "" Then
		$g_oCache[$sMemKey] = $aInfo[1]
		Return $aInfo[1]
	EndIf

	; Вызов ffmpeg.exe -i (парсинг разрешения из строки Video:)
	Local $sCmdLine = '"' & $g_sPathFFmpeg & '" -hide_banner -i "' & $sVideoPath & '"'
	Local $sRaw = _RunToolReadStderr($sCmdLine)
	Local $aMatch = StringRegExp($sRaw, "Video:\s.*?,\s(\d+)x(\d+)", 1)
	Local $sOutput = (IsArray($aMatch) ? $aMatch[0] & "," & $aMatch[1] : "")
	ConsoleWrite($sVideoPath & " -> " & $sOutput & @CRLF)

	; Сохраняем в оба кеша
	If $sOutput <> "" Then
		$g_oCache[$sMemKey] = $sOutput
		_SaveCacheInfo($sVideoPath, $sOutput)
	EndIf

	Return $sOutput
EndFunc   ;==>_GetVideoResolution


Func _RunToolReadStderr($sCmdLine)
	Local $iPid = Run($sCmdLine, "", @SW_HIDE, $STDERR_CHILD + $STDOUT_CHILD)
	Local $sOutput = "", $sLine
	While 1
		$sLine = StderrRead($iPid)
		If @error Then ExitLoop
		$sOutput &= $sLine
	WEnd
	Return StringStripWS($sOutput, 3)
EndFunc   ;==>_RunToolReadStderr


Func _SetCtrlResizing()

	; Комбинации флагов GUICtrlSetResizing
	Local $iDockFixed = $GUI_DOCKLEFT + $GUI_DOCKTOP + $GUI_DOCKWIDTH + $GUI_DOCKHEIGHT
	Local $iDockStretchH = $GUI_DOCKLEFT + $GUI_DOCKTOP + $GUI_DOCKRIGHT + $GUI_DOCKHEIGHT
	Local $iDockFixedRight = $GUI_DOCKRIGHT + $GUI_DOCKTOP + $GUI_DOCKWIDTH + $GUI_DOCKHEIGHT
	Local $iDockFixedBL = $GUI_DOCKLEFT + $GUI_DOCKBOTTOM + $GUI_DOCKWIDTH + $GUI_DOCKHEIGHT
	Local $iDockFixedBR = $GUI_DOCKRIGHT + $GUI_DOCKBOTTOM + $GUI_DOCKWIDTH + $GUI_DOCKHEIGHT
	Local $iDockStretchHV = $GUI_DOCKLEFT + $GUI_DOCKTOP + $GUI_DOCKRIGHT + $GUI_DOCKBOTTOM
	Local $iDockStretchH_B = $GUI_DOCKLEFT + $GUI_DOCKBOTTOM + $GUI_DOCKRIGHT + $GUI_DOCKHEIGHT

	; Растягиваемые по горизонтали
	GUICtrlSetResizing($g_iLogoPic, $iDockFixed)
	GUICtrlSetResizing($g_iInput1, $iDockStretchH)
	GUICtrlSetResizing($g_iLabelInfo1, $iDockStretchH)
	GUICtrlSetResizing($g_iInput2, $iDockStretchH)
	GUICtrlSetResizing($g_iLabelInfo2, $iDockStretchH)
	GUICtrlSetResizing($g_iSeparator, $iDockStretchH)
	GUICtrlSetResizing($g_iSeparatorTop, $iDockStretchH)
	GUICtrlSetResizing($g_iEditCommand, $iDockStretchHV)

	; Фиксированные слева
	GUICtrlSetResizing($g_iLabel1, $iDockFixed)
	GUICtrlSetResizing($g_iLabelOffset, $iDockFixed)
	GUICtrlSetResizing($g_iRadioOffsetAuto, $iDockFixed)
	GUICtrlSetResizing($g_iLabelSyncStatus, $iDockFixed)
	GUICtrlSetResizing($g_iLabel2, $iDockFixed)
	GUICtrlSetResizing($g_iLabelCompare, $iDockFixed)
	GUICtrlSetResizing($g_iRadioDirect, $iDockFixed)
	GUICtrlSetResizing($g_iRadioVertical, $iDockFixed)
	GUICtrlSetResizing($g_iLabelCommand, $iDockFixed)

	; Фиксированные справа
	GUICtrlSetResizing($g_iButtonSettings, $iDockFixedRight)
	GUICtrlSetResizing($g_iButtonChoose1, $iDockFixedRight)
	GUICtrlSetResizing($g_iButtonSwap, $iDockFixedRight)
	GUICtrlSetResizing($g_iButtonChoose2, $iDockFixedRight)
	GUICtrlSetResizing($g_iButtonCompare, $iDockFixedBR)
	GUICtrlSetResizing($g_iRadioOffsetManual, $iDockFixedRight)
	GUICtrlSetResizing($g_iInputOffset, $iDockFixedRight)
EndFunc   ;==>_SetCtrlResizing



Func _RenderHeader()
	; Новый порядок: help, info, HUD, mode, toggle, seek1, seek15, shift10, shift100, shift1 (нижняя первой колонки — наверх второй)
	Local $aHintRows[10]
	$aHintRows[0] = Lang("Hotkeys", "2", "{H} Control hints")
	$aHintRows[1] = Lang("Hotkeys", "3", "{Y} Video info")
	$aHintRows[2] = Lang("Hotkeys", "7", "{3} Show/hide HUD")
	$aHintRows[3] = Lang("Hotkeys", "8", "{0} Video/subtraction mode")
	$aHintRows[4] = Lang("Hotkeys", "9", "{Y} Toggle subtraction mode")
	$aHintRows[5] = Lang("Hotkeys", "4", "{LEFT}{RIGHT} Seek ±1 sec")
	$aHintRows[6] = Lang("Hotkeys", "5", "{UP}{DOWN} Seek ±15 sec")
	$aHintRows[7] = Lang("Hotkeys", "10", "{PLUS}{MINUS} Shift ±1 frame")
	$aHintRows[8] = Lang("Hotkeys", "11", "{CTRL}{PLUS}{MINUS} Shift ±10 frames")
	$aHintRows[9] = Lang("Hotkeys", "12", "{ALT}{PLUS}{MINUS} Shift ±100 frames")


	Local $sIconsRoot = @ScriptDir & "\Assets\KeyIcons"
	Local $sKeyTheme = $g_bDarkMode ? "Dark" : "Light"
	Local $sHeaderIcon = @ScriptDir & "\Assets\Icons\HeaderIcon.png"

	; Заголовки шапки под иконкой: VCLauncher, затем Video-compare и его версия отдельной строкой.
	Local $sTitleApp = $gc_sAppName
	Local $sTitleVc = "Video-compare"
	Local $sTitleVcVer = $g_sVcVersion
	Local $sTitleRight = Lang("Hotkeys", "1", "Video-compare hotkeys")

	; Фон шапки = фон GUI, чтобы полоса справа под кнопкой не отличалась по тону.
	; Берём от $g_bDarkMode, т.к. _SetPalette может ещё не отработать при первом
	; рендере на старте. ARGB = непрозрачный фон GUI.
	Local $iBackArgb = 0xFF000000 + ($g_bDarkMode ? $gc_iClrBgDark : $gc_iClrBgLight)

	_HeaderRenderToPic($g_iLogoPic, $g_hLogoBitmap, $sIconsRoot, $sKeyTheme, $g_sAppFont, $gc_iHeaderW, $gc_iHeaderH, $aHintRows, 0, 9, 0, $iBackArgb, $sHeaderIcon, $sTitleApp, $sTitleVc, $sTitleRight, $sTitleVcVer)
EndFunc   ;==>_RenderHeader


Func _SetButtonIcon($iCtrl, $sDllPath, $iIconIndex, $iIconSize)
	; Извлекаем иконку нужного размера
	Local $aIcons = DllCall("user32.dll", "uint", "PrivateExtractIconsW", _
			"wstr", $sDllPath, "int", $iIconIndex, "int", $iIconSize, "int", $iIconSize, _
			"handle*", 0, "uint*", 0, "uint", 1, "uint", 0)
	If @error Or $aIcons[0] = 0 Then Return 0
	Return _ApplyHIconToButton($iCtrl, $aIcons[5], $iIconSize)
EndFunc   ;==>_SetButtonIcon


; Оборачивает HICON в новый HIMAGELIST и назначает его кнопке.
; HICON уничтожается после добавления в имидж-лист. Возвращает HIMAGELIST
; (вызывающий ответственен за ImageList_Destroy старого списка при замене).
; $iAlign: 2 = TOP (иконка над текстом), 4 = CENTER (только иконка). $iMarginTop — верхний отступ.
Func _ApplyHIconToButton($iCtrl, $hIcon, $iIconSize, $iAlign = 2, $iMarginTop = 10)
	Local $aImgList = DllCall("comctl32.dll", "handle", "ImageList_Create", _
			"int", $iIconSize, "int", $iIconSize, "uint", $ILC_COLOR32, "int", 1, "int", 0)
	If @error Then
		DllCall("user32.dll", "bool", "DestroyIcon", "handle", $hIcon)
		Return 0
	EndIf
	Local $hImgList = $aImgList[0]

	DllCall("comctl32.dll", "int", "ImageList_ReplaceIcon", _
			"handle", $hImgList, "int", -1, "handle", $hIcon)
	DllCall("user32.dll", "bool", "DestroyIcon", "handle", $hIcon)

	; BUTTON_IMAGELIST: himl, margin(l,t,r,b), uAlign
	Local $tBIL = DllStructCreate("handle himl;int l;int t;int r;int b;uint uAlign")
	DllStructSetData($tBIL, "himl", $hImgList)
	DllStructSetData($tBIL, "l", 0)
	DllStructSetData($tBIL, "t", $iMarginTop)
	DllStructSetData($tBIL, "r", 0)
	DllStructSetData($tBIL, "b", 0)
	DllStructSetData($tBIL, "uAlign", $iAlign)

	DllCall("user32.dll", "lresult", "SendMessageW", "hwnd", GUICtrlGetHandle($iCtrl), _
			"uint", $BCM_SETIMAGELIST, "wparam", 0, "struct*", $tBIL)

	Return $hImgList
EndFunc   ;==>_ApplyHIconToButton


; Цвет иконок = цвет текста кнопки текущей темы, чуть мягче (светлая — тёмно-серый, тёмная — светлый).
Func _IconTint()
	Return $g_bDarkMode ? 0xCCCCCC : 0x303030
EndFunc   ;==>_IconTint


; Ставит PNG-иконку из Assets\Icons по центру кнопки, перекрашенную под тему. Возвращает HIMAGELIST.
; $iTint = -1 — цвет по теме (_IconTint); иначе явный 0xRRGGBB (для пульсации кнопки).
Func _SetButtonCenterIcon($iCtrl, $sPngName, $iIconSize, $iTint = -1)
	If $iTint = -1 Then $iTint = _IconTint()
	Local $sPath = @ScriptDir & "\Assets\Icons\" & $sPngName
	Local $hIcon = _LoadHIconFromPng($sPath, $iIconSize, $iTint)
	If Not $hIcon Then Return 0
	Return _ApplyHIconToButton($iCtrl, $hIcon, $iIconSize, 4, 0) ; 4 = BUTTON_IMAGELIST_ALIGN_CENTER
EndFunc   ;==>_SetButtonCenterIcon


; (Пере)назначает иконки кнопок под текущую тему. Освобождает старые imagelist'ы.
Func _ApplyButtonIcons()
	_DestroyImgList($g_hSettingsImgList)
	_DestroyImgList($g_hSwapImgList)
	$g_hSettingsImgList = _SetButtonCenterIcon($g_iButtonSettings, "Settings.png", 18)
	$g_hSwapImgList = _SetButtonCenterIcon($g_iButtonSwap, "Swap.png", 16)
	_UpdateCompareButtonIcon() ; сама освобождает старый $g_hCompareImgList
EndFunc   ;==>_ApplyButtonIcons


; Загружает PNG в HICON квадратного размера $iIconSize через GDI+.
; $iTint >= 0 — перекрашивает иконку в цвет 0xRRGGBB (силуэт по альфе), сохраняя сглаживание.
Func _LoadHIconFromPng($sPngPath, $iIconSize, $iTint = -1)
	If Not FileExists($sPngPath) Then Return 0

	_GDIPlus_Startup()
	Local $hSrc = _GDIPlus_ImageLoadFromFile($sPngPath)
	If @error Or Not $hSrc Then
		_GDIPlus_Shutdown()
		Return 0
	EndIf

	Local $hScaled = _GDIPlus_ImageResize($hSrc, $iIconSize, $iIconSize)
	_GDIPlus_ImageDispose($hSrc)
	If @error Or Not $hScaled Then
		_GDIPlus_Shutdown()
		Return 0
	EndIf

	If $iTint >= 0 Then
		Local $hTinted = __TintCopy($hScaled, $iIconSize, $iTint)
		If $hTinted Then
			_GDIPlus_BitmapDispose($hScaled)
			$hScaled = $hTinted
		EndIf
	EndIf

	Local $hIcon = _GDIPlus_HICONCreateFromBitmap($hScaled)
	_GDIPlus_BitmapDispose($hScaled)
	_GDIPlus_Shutdown()

	If @error Or Not $hIcon Then Return 0
	Return $hIcon
EndFunc   ;==>_LoadHIconFromPng


; Возвращает копию GDI+ bitmap, перекрашенную в сплошной цвет $iColor (0xRRGGBB),
; с сохранением альфы (силуэт). Требует уже запущенного GDI+. Возвращает 0 при ошибке.
Func __TintCopy($hImg, $iSize, $iColor)
	Local $hBmp = _GDIPlus_BitmapCreateFromScan0($iSize, $iSize)
	If @error Or Not $hBmp Then Return 0
	Local $hGfx = _GDIPlus_ImageGetGraphicsContext($hBmp)
	_GDIPlus_GraphicsSetInterpolationMode($hGfx, 7) ; HighQualityBicubic

	Local $nR = BitShift(BitAND($iColor, 0xFF0000), 16) / 255
	Local $nG = BitShift(BitAND($iColor, 0x00FF00), 8) / 255
	Local $nB = BitAND($iColor, 0x0000FF) / 255

	; Обнуляем вклад входных RGB, альфу пропускаем, RGB задаём смещением (offset-строка)
	Local $tCM = DllStructCreate("float[25]")
	DllStructSetData($tCM, 1, 1.0, 19) ; A -> A
	DllStructSetData($tCM, 1, $nR, 21) ; +R
	DllStructSetData($tCM, 1, $nG, 22) ; +G
	DllStructSetData($tCM, 1, $nB, 23) ; +B
	DllStructSetData($tCM, 1, 1.0, 25)

	Local $aAttr = DllCall("gdiplus.dll", "int", "GdipCreateImageAttributes", "ptr*", 0)
	If @error Or Not IsArray($aAttr) Or $aAttr[0] <> 0 Then
		_GDIPlus_GraphicsDispose($hGfx)
		_GDIPlus_BitmapDispose($hBmp)
		Return 0
	EndIf
	Local $hAttr = $aAttr[1]
	Local Const $ColorAdjustTypeBitmap = 1
	DllCall("gdiplus.dll", "int", "GdipSetImageAttributesColorMatrix", _
			"ptr", $hAttr, "int", $ColorAdjustTypeBitmap, "bool", True, _
			"ptr", DllStructGetPtr($tCM), "ptr", 0, "int", 0)

	Local Const $UnitPixel = 2
	DllCall("gdiplus.dll", "int", "GdipDrawImageRectRectI", _
			"handle", $hGfx, "handle", $hImg, _
			"int", 0, "int", 0, "int", $iSize, "int", $iSize, _
			"int", 0, "int", 0, "int", $iSize, "int", $iSize, _
			"int", $UnitPixel, "ptr", $hAttr, "ptr", 0, "ptr", 0)

	DllCall("gdiplus.dll", "int", "GdipDisposeImageAttributes", "ptr", $hAttr)
	_GDIPlus_GraphicsDispose($hGfx)
	Return $hBmp
EndFunc   ;==>__TintCopy


; Устанавливает иконку кнопки «Сравнить» согласно выбранному режиму радио.
Func _UpdateCompareButtonIcon()
	Local $bIsVertical = (GUICtrlRead($g_iRadioVertical) = $GUI_CHECKED)
	Local $sIconName = $bIsVertical ? "CompareVstack.png" : "CompareDirect.png"
	Local $sPath = @ScriptDir & "\Assets\Icons\" & $sIconName
	Local Const $iIconSize = 32

	Local $hIcon = _LoadHIconFromPng($sPath, $iIconSize, _IconTint()) ; перекраска под тему
	If Not $hIcon Then Return

	Local $hNew = _ApplyHIconToButton($g_iButtonCompare, $hIcon, $iIconSize)
	If Not $hNew Then Return

	_DestroyImgList($g_hCompareImgList) ; освобождаем предыдущий imagelist после замены
	$g_hCompareImgList = $hNew
EndFunc   ;==>_UpdateCompareButtonIcon


Func _RestoreControlsStyle()
	AdlibUnRegister("_RestoreControlsStyle")

	; Снимаем drag-подсветку: сбрасываем флаг и перерисовываем поле обычным цветом UDF
	Local $hPrev = $g_hDragHiCtrl
	$g_hDragHiCtrl = 0
	If $hPrev Then _WinAPI_InvalidateRect($hPrev)

	; Обновляем информацию о файлах
	_UpdateFilesInfo()
EndFunc   ;==>_RestoreControlsStyle


Func _GetVideoExtensionsFilter()
	Local $sExtensions = ""
	Local $sExt = ''
	For $sExt In $gc_aSupportedExtensions
		$sExtensions &= "*." & $sExt & ";"
	Next
	Return Lang("Filter", "Video", "Video") & " (" & StringTrimRight($sExtensions, 1) & ")"
EndFunc   ;==>_GetVideoExtensionsFilter


Func _SetComboItemHeight($iCtrl, $iItemHeight)
	Local $hWnd = GUICtrlGetHandle($iCtrl)
	If Not $hWnd Then Return
	; wParam = -1 → «поле выбора» комбобокса (selection field)
	DllCall("user32.dll", "lresult", "SendMessageW", "hwnd", $hWnd, "uint", $CB_SETITEMHEIGHT, "wparam", -1, "lparam", $iItemHeight)
EndFunc   ;==>_SetComboItemHeight


Func _ApplyAppFont()
	$g_sAppFont = _FontApply($g_hGui, $g_iAppFontSize)
EndFunc   ;==>_ApplyAppFont


Func _PathGetDir($sPath)
	Local $iPos = StringInStr($sPath, "\", 0, -1)
	If $iPos > 0 Then
		Return StringLeft($sPath, $iPos - 1)
	EndIf
	Return ""
EndFunc   ;==>_PathGetDir


Func _GetFileName($sPath)
	Return StringRegExpReplace($sPath, '^.*[\\/]', '')
EndFunc   ;==>_GetFileName


Func _ResetInputCaret($iCtrlID)
	GUICtrlSendMsg($iCtrlID, $EM_SETSEL, 0, 0)
EndFunc   ;==>_ResetInputCaret


Func _IsValidVideoExtension($sFilePath)
	Local $sExt = StringLower(StringRegExpReplace($sFilePath, '^.*\.', ''))
	For $sSupportedExt In $gc_aSupportedExtensions
		If $sExt = $sSupportedExt Then Return True
	Next
	Return False
EndFunc   ;==>_IsValidVideoExtension


Func _SetVideoFile($iInputID, $sFilePath)
	If $iInputID = $g_iInput1 Then
		$g_sVideoFile1 = $sFilePath
		GUICtrlSetData($g_iInput1, $sFilePath = "" ? "" : _GetFileName($sFilePath))
		_ResetInputCaret($g_iInput1)
		IniWrite($gc_sPathIni, "LastDirs", "Video1", $sFilePath)
	ElseIf $iInputID = $g_iInput2 Then
		$g_sVideoFile2 = $sFilePath
		GUICtrlSetData($g_iInput2, $sFilePath = "" ? "" : _GetFileName($sFilePath))
		_ResetInputCaret($g_iInput2)
		IniWrite($gc_sPathIni, "LastDirs", "Video2", $sFilePath)
	EndIf
	GUICtrlSetData($g_iInputOffset, "")
	_TryFillCachedOffset()
	_UpdateFilesInfo()
EndFunc   ;==>_SetVideoFile


Func _EnsureIniDefaults()
	If FileExists($gc_sPathIni) Then Return
	; Создаём ini с UTF-16 LE BOM (поддержка Unicode-путей) и значениями по умолчанию
	Local $sContent = "[LastDirs]" & @CRLF & _
			"Video1=" & @CRLF & _
			"Video2=" & @CRLF & _
			"[Tools]" & @CRLF & _
			"VideoCompare=video-compare.exe" & @CRLF & _
			"Sync=Sync\dist\Sync.exe" & @CRLF & _
			"FFmpeg=Sync\dist\FFmpeg.exe" & @CRLF & _
			"[Settings]" & @CRLF & _
			"SyncSkipSec=300" & @CRLF & _
			"SyncTimeoutSec=60" & @CRLF & _
			"SyncMethod=audio" & @CRLF & _
			"VcVersion=" & @CRLF & _
			"UpdateCheckFreq=weekly" & @CRLF & _
			"LastUpdateCheck=" & @CRLF
	_EnsureUtf16File($gc_sPathIni, $sContent)
EndFunc   ;==>_EnsureIniDefaults


; Создаёт пустой файл с UTF-16 LE BOM, если он ещё не существует
Func _EnsureUtf16File($sFilePath, $sContent = "")
	If FileExists($sFilePath) Then Return
	Local $hFile = FileOpen($sFilePath, $FO_OVERWRITE + $FO_UTF16_LE)
	If $hFile = -1 Then Return
	If $sContent <> "" Then FileWrite($hFile, $sContent)
	FileClose($hFile)
EndFunc   ;==>_EnsureUtf16File


Func _NormalizePath($sValue)
	; Пустая строка — возвращаем как есть, иначе превратится в путь к папке
	If $sValue = "" Then Return ""
	; Абсолютный путь: диск:\ или UNC \\ или корень /\
	If StringRegExp($sValue, "^(?:[A-Za-z]:\\|\\\\|/)") Then Return $sValue
	; Иначе считаем относительным к папке скрипта
	Return @ScriptDir & "\" & $sValue
EndFunc   ;==>_NormalizePath


; Проверка наличия инструмента: при отсутствии — сообщение и открытие ini
Func _CheckToolExists($sToolPath, $sToolName)
	If FileExists($sToolPath) Then Return
	MsgBox(48, $gc_sAppName, StringReplace(Lang("Errors", "ToolNotFound", '"%TOOL%" not found.'), "%TOOL%", $sToolName) & @CR & @CR & _
			Lang("Errors", "SettingsWillOpen", "The settings file will now be opened."))
	ShellExecute($gc_sPathIni)
	Exit
EndFunc   ;==>_CheckToolExists


Func _ResolveToolPaths()
	Local $sIniVideoCompare = _NormalizePath(IniRead($gc_sPathIni, "Tools", "VideoCompare", ""))
	If Not FileExists($g_sPathVideoCompare) And FileExists($sIniVideoCompare) Then
		$g_sPathVideoCompare = $sIniVideoCompare
	EndIf

	Local $sIniSync = _NormalizePath(IniRead($gc_sPathIni, "Tools", "Sync", ""))
	If Not FileExists($g_sPathSync) And FileExists($sIniSync) Then
		$g_sPathSync = $sIniSync
	EndIf

	Local $sIniFFmpeg = _NormalizePath(IniRead($gc_sPathIni, "Tools", "FFmpeg", ""))
	If Not FileExists($g_sPathFFmpeg) And FileExists($sIniFFmpeg) Then
		$g_sPathFFmpeg = $sIniFFmpeg
	EndIf
EndFunc   ;==>_ResolveToolPaths


; Версия video-compare для шапки. Берётся из ini; если там пусто — определяется
; через --version и кешируется в настройки. Обновление происходит только когда
; версия в ini не указана (по требованию).
Func _InitVcVersion()
	$g_sVcVersion = IniRead($gc_sPathIni, "Settings", "VcVersion", "")
	If $g_sVcVersion <> "" Then Return
	$g_sVcVersion = _DetectVcVersion()
	If $g_sVcVersion <> "" Then IniWrite($gc_sPathIni, "Settings", "VcVersion", $g_sVcVersion)
EndFunc   ;==>_InitVcVersion


; Запрашивает версию у video-compare через "--version".
; Вывод формата "video-compare 20260502-valencia" приводим к виду '20260502 "valencia"'.
; Возвращает "" при ошибке запуска или нераспознанном выводе.
Func _DetectVcVersion()
	Local $iPid = Run('"' & $g_sPathVideoCompare & '" --version', @ScriptDir, @SW_HIDE, $STDOUT_CHILD)
	If @error Or Not $iPid Then Return ""
	ProcessWaitClose($iPid, 5)
	Local $sOut = StdoutRead($iPid)
	; Токен версии после "video-compare"
	Local $aMatch = StringRegExp($sOut, '(?im)video-compare\s+(\S+)', 1)
	If @error Then Return ""
	Local $sRaw = $aMatch[0] ; напр. 20260502-valencia
	; Дату и кодовое имя в кавычках разделяем по первому дефису
	Local $iDash = StringInStr($sRaw, "-")
	If $iDash > 0 Then Return StringLeft($sRaw, $iDash - 1) & ' "' & StringTrimLeft($sRaw, $iDash) & '"'
	Return $sRaw
EndFunc   ;==>_DetectVcVersion


; === Система локализации ===

Func Lang($sSection, $sKey, $sDefault = "")
	Local $sFullKey = $sSection & "." & $sKey
	If MapExists($g_mLang, $sFullKey) Then Return $g_mLang[$sFullKey]
	Return $sDefault
EndFunc   ;==>Lang


Func _InitLanguage()
	$g_sCurrentLang = IniRead($gc_sPathIni, "Settings", "Language", "")
	If $g_sCurrentLang = "" Then
		; Авто-определение по языку системы
		If @OSLang = "0419" Or @OSLang = "0422" Then
			$g_sCurrentLang = "Russian"
		Else
			$g_sCurrentLang = "English"
		EndIf
	EndIf
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
	Local $aLines = StringSplit(StringReplace($sContent, @CRLF, @LF), @LF)
	Local $sSection = ""
	For $i = 1 To $aLines[0]
		Local $sLine = StringStripWS($aLines[$i], 3)
		If $sLine = "" Or StringLeft($sLine, 1) = ";" Then ContinueLoop
		If StringLeft($sLine, 1) = "[" And StringRight($sLine, 1) = "]" Then
			$sSection = StringMid($sLine, 2, StringLen($sLine) - 2)
			ContinueLoop
		EndIf
		Local $iEq = StringInStr($sLine, "=")
		If $iEq > 0 And $sSection <> "" Then
			Local $sK = StringStripWS(StringLeft($sLine, $iEq - 1), 3)
			Local $sV = StringMid($sLine, $iEq + 1)
			$g_mLang[$sSection & "." & $sK] = $sV
		EndIf
	Next
EndFunc   ;==>_LoadLangFile


Func _ApplyLanguage()
	GUICtrlSetData($g_iLabel1, Lang("GUI", "File1", "File 1"))
	GUICtrlSetData($g_iLabel2, Lang("GUI", "File2", "File 2"))
	GUICtrlSetData($g_iButtonChoose1, Lang("GUI", "Choose", "Choose"))
	GUICtrlSetData($g_iButtonChoose2, Lang("GUI", "Choose", "Choose"))
	GUICtrlSetData($g_iLabelCompare, Lang("GUI", "CompareMode", "Compare:"))
	GUICtrlSetData($g_iRadioDirect, Lang("GUI", "CompareDirect", "Direct"))
	GUICtrlSetData($g_iRadioVertical, Lang("GUI", "CompareVertical", "Vertical"))
	GUICtrlSetData($g_iButtonCompare, Lang("GUI", "Compare", "Compare"))
	GUICtrlSetData($g_iLabelOffset, Lang("GUI", "Offset", "Offset:"))
	GUICtrlSetData($g_iRadioOffsetAuto, Lang("GUI", "OffsetAuto", "Auto"))
	GUICtrlSetData($g_iRadioOffsetManual, Lang("GUI", "OffsetManual", "Manual"))
	GUICtrlSetData($g_iLabelCommand, Lang("GUI", "TabCommand", "Command"))
	_RenderHeader()
	_UpdateFilesInfo()
	_ApplyTheme()
	_ApplyUpdateNotice() ; перевести постфикс «[Доступно обновление]» в заголовке
EndFunc   ;==>_ApplyLanguage


; === Темы оформления ===

Func _InitTheme()
	$g_sTheme = IniRead($gc_sPathIni, "Settings", "Theme", "System")
	If $g_sTheme <> "System" And $g_sTheme <> "Light" And $g_sTheme <> "Dark" Then
		$g_sTheme = "System"
	EndIf
	_ResolveDarkMode()
EndFunc   ;==>_InitTheme


Func _ResolveDarkMode()
	Switch $g_sTheme
		Case "Light"
			$g_bDarkMode = False
		Case "Dark"
			$g_bDarkMode = True
		Case Else
			$g_bDarkMode = _ReadSystemDarkMode()
	EndSwitch
EndFunc   ;==>_ResolveDarkMode


Func _ReadSystemDarkMode()
	Local $vVal = RegRead("HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Themes\Personalize", "AppsUseLightTheme")
	If @error Then Return False
	Return ($vVal = 0)
EndFunc   ;==>_ReadSystemDarkMode


Func _SetPalette()
	If $g_bDarkMode Then
		$g_iClrBg = $gc_iClrBgDark
		$g_iClrFg = $COLOR_TEXT_LIGHT
		$g_iClrInfo = 0x969AA2 ; как подзаголовок шапки (HeaderHelper $iColSub, dark)
		$g_iClrSep = $COLOR_BORDER
	Else
		$g_iClrBg = $gc_iClrBgLight
		$g_iClrFg = 0x000000
		$g_iClrInfo = $gc_iClrInfoLight
		$g_iClrSep = $gc_iClrSepLight
	EndIf
EndFunc   ;==>_SetPalette


Func _ApplyTheme()
	_ResolveDarkMode()

	; Сравниваем с реально применённой к GUI темой, а не с $g_bDarkMode до вызова:
	; _InitTheme() может быть вызван выше по стеку и уже обновить $g_bDarkMode.
	Local $bNeedSwitch = (Not $g_bThemeInitialized) Or ($g_bAppliedDark <> $g_bDarkMode)

	; Применяем UDF-тему ко всему GUI. Не используем _GUIDarkTheme_SwitchTheme —
	; он определяет направление по системной теме, а не по нашему выбору.
	; Перед сменой темы задаём проектную палитру ($COLOR_*, $g_iBkColor) для кастомной
	; отрисовки. UDF свои внутренние цвета сбрасывает сам при каждом GUISetDarkTheme.
	If $bNeedSwitch Then
		; Полная очистка перед сменой темы. UDF кэширует кисти/перья (guard "If Not $h"),
		; поэтому без BrushCleanup/PenCleanup они остаются старого цвета и поля красятся
		; неверно. _GUIDarkTheme_SwitchTheme делает все три cleanup; повторяем то же.
		If $g_bThemeInitialized Then
			__GUIDarkTheme_SubclassCleanup()
			__GUIDarkTheme_BrushCleanup()
			__GUIDarkTheme_PenCleanup()
		EndIf
		If $g_bDarkMode Then
			$g_iBkColor = $gc_iClrBgDark
			$COLOR_BG_DARK = $gc_iClrBgDark
			$COLOR_TEXT_LIGHT = 0xCCCCCC
			$COLOR_CONTROL_BG = 0x2D2D2D
			$COLOR_BORDER_LIGHT = 0x555555
			$COLOR_BORDER = 0x3C3C3C
			_GUIDarkTheme_ApplyDark($g_hGui)
		Else
			_GUIDarkTheme_ApplyLight($g_hGui)
		EndIf
		$g_bThemeInitialized = True
		$g_bAppliedDark = $g_bDarkMode
		; UDF при ApplyDark/Light переустанавливает свой WM_CTLCOLOREDIT-хендлер.
		; Перебиваем его нашим делегатом, иначе drag-подсветка не отрисуется.
		GUIRegisterMsg($WM_CTLCOLOREDIT, "_OnEvent_WM_CTLCOLOREDIT")
	EndIf

	; Палитра для кастомных элементов
	_SetPalette()

	; UDF ставит фон окна 0x121212 — в тёмной теме делаем его светлее
	If $g_bDarkMode Then GUISetBkColor($g_iClrBg, $g_hGui)

	; Статические лейблы тела окна. UDF в кейсе "Static" делает им ПРОЗРАЧНЫЙ фон
	; (GUICtrlSetBkColor TRANSPARENT) и красит текст. У прозрачного static текст при каждой
	; перерисовке накладывается на старый — RDW_ERASE родителя не чистит чужой HWND,
	; поэтому шрифт «жирнеет» с каждой сменой темы. Ставим НЕПРОЗРАЧНЫЙ фон = цвет тела окна:
	; $g_iClrBg точно совпадает с фоном (light 0xF0F0F0, dark 0x1E1E1E через GUISetBkColor),
	; вид тот же, но static стирает фон сплошной кистью и накопления нет.
	Local $aBodyLabels[5] = [$g_iLabel1, $g_iLabel2, $g_iLabelCompare, $g_iLabelOffset, $g_iLabelCommand]
	For $iLbl In $aBodyLabels
		GUICtrlSetBkColor($iLbl, $g_iClrBg)
		GUICtrlSetColor($iLbl, $g_iClrFg)
	Next

	; Info-лейблы — серый оттенок, непрозрачный фон по той же причине
	GUICtrlSetBkColor($g_iLabelInfo1, $g_iClrBg)
	GUICtrlSetBkColor($g_iLabelInfo2, $g_iClrBg)
	GUICtrlSetColor($g_iLabelInfo1, $g_iClrInfo)
	GUICtrlSetColor($g_iLabelInfo2, $g_iClrInfo)

	; Поле сдвига красит UDF (edit), проектная покраска убрана как дубль.

	; Статус sync — непрозрачный фон тела. Цвет символа задаётся в _SetSyncStatus.
	GUICtrlSetBkColor($g_iLabelSyncStatus, $g_iClrBg)
	; При смене темы пересчитываем цвет под текущий статус
	If $g_sLastSyncStatus <> "" Then _SetSyncStatus($g_sLastSyncStatus, $g_iLastSyncOffset)

	; Иконки кнопок перекрашиваем под цвет текста темы
	_ApplyButtonIcons()

	; Горизонтальный разделитель
	GUICtrlSetBkColor($g_iSeparator, $g_iClrSep)
	GUICtrlSetBkColor($g_iSeparatorTop, $g_iClrSep)

	; Иконки подсказок должны соответствовать выбранной теме (Dark/Light)
	If $bNeedSwitch Then _RenderHeader()

	; Финальная полная перерисовка окна. RDW_ERASE обязателен: UDF переводит static-лейблы
	; в прозрачный фон ($GUI_BKCOLOR_TRANSPARENT), и без стирания фона текст накладывается
	; на старый при каждой смене темы (лейблы выглядят всё жирнее).
	DllCall("user32.dll", "bool", "RedrawWindow", "hwnd", $g_hGui, "ptr", 0, "handle", 0, "uint", $RDW_INVALIDATE + $RDW_ERASE + $RDW_UPDATENOW + $RDW_ALLCHILDREN)
EndFunc   ;==>_ApplyTheme




Func _SettingsWindow()
	; По центру основного окна (а не экрана)
	Local Const $iSetW = 302, $iSetH = 196 ; отступы 16px с обеих сторон (правый край контента = 286)
	Local $iSetX = -1, $iSetY = -1
	Local $aMain = WinGetPos($g_hGui)
	If Not @error Then
		$iSetX = $aMain[0] + Int(($aMain[2] - $iSetW) / 2)
		$iSetY = $aMain[1] + Int(($aMain[3] - $iSetH) / 2)
	EndIf
	$g_hSettingsGui = GUICreate(Lang("GUI", "Settings", "Settings"), $iSetW, $iSetH, $iSetX, $iSetY, _
			$WS_CAPTION + $WS_SYSMENU, $WS_EX_DLGMODALFRAME, $g_hGui)

	Local $iY = 20

	; Язык
	$g_iSettingsLabelLang = GUICtrlCreateLabel(Lang("GUI", "Language", "Language:"), 16, $iY + 3, 75, 20)
	$g_iSettingsComboLang = GUICtrlCreateCombo("", 96, $iY, 190, 200, $CBS_DROPDOWNLIST)
	GUICtrlSetData($g_iSettingsComboLang, "English|Русский", ($g_sCurrentLang = "Russian") ? "Русский" : "English")
	_SetComboItemHeight($g_iSettingsComboLang, 17)

	$iY += 38

	; Тема
	$g_iSettingsLabelTheme = GUICtrlCreateLabel(Lang("GUI", "Theme", "Theme:"), 16, $iY + 3, 75, 20)
	$g_iSettingsComboTheme = GUICtrlCreateCombo("", 96, $iY, 190, 200, $CBS_DROPDOWNLIST)
	Local $sSystem = Lang("GUI", "ThemeSystem", "System")
	Local $sLight = Lang("GUI", "ThemeLight", "Light")
	Local $sDark = Lang("GUI", "ThemeDark", "Dark")
	Local $sCurrent = $sSystem
	If $g_sTheme = "Light" Then $sCurrent = $sLight
	If $g_sTheme = "Dark" Then $sCurrent = $sDark
	GUICtrlSetData($g_iSettingsComboTheme, $sSystem & "|" & $sLight & "|" & $sDark, $sCurrent)
	_SetComboItemHeight($g_iSettingsComboTheme, 17)

	$iY += 38

	; Обновления: подпись + combo + квадратная кнопка ручной проверки.
	; Порядок пунктов: Никогда / При запуске / Раз в неделю / Раз в месяц
	$g_iSettingsLabelUpdates = GUICtrlCreateLabel(Lang("Updates", "CheckLabel", "Updates:"), 16, $iY + 3, 75, 20)
	$g_iSettingsComboUpdates = GUICtrlCreateCombo("", 96, $iY, 162, 200, $CBS_DROPDOWNLIST)
	Local $sUpdNever   = Lang("Updates", "FreqNever", "Never")
	Local $sUpdAlways  = Lang("Updates", "FreqAlways", "At startup")
	Local $sUpdWeekly  = Lang("Updates", "FreqWeekly", "Once a week")
	Local $sUpdMonthly = Lang("Updates", "FreqMonthly", "Once a month")
	Local $sFreq = IniRead($gc_sPathIni, "Settings", "UpdateCheckFreq", "weekly")
	Local $sFreqCur = $sUpdWeekly
	Switch $sFreq
		Case "never"
			$sFreqCur = $sUpdNever
		Case "always"
			$sFreqCur = $sUpdAlways
		Case "monthly"
			$sFreqCur = $sUpdMonthly
	EndSwitch
	GUICtrlSetData($g_iSettingsComboUpdates, $sUpdNever & "|" & $sUpdAlways & "|" & $sUpdWeekly & "|" & $sUpdMonthly, $sFreqCur)
	_SetComboItemHeight($g_iSettingsComboUpdates, 17)
	GUICtrlSetOnEvent($g_iSettingsComboUpdates, "_OnEvent_SettingsUpdatesCombo")

	; Квадратная кнопка проверки (иконка круга со стрелками назначается после темы)
	$g_iSettingsButtonCheckNow = GUICtrlCreateButton("", 262, $iY - 1, 24, 23)
	GUICtrlSetTip($g_iSettingsButtonCheckNow, Lang("Updates", "CheckNow", "Check now"))
	GUICtrlSetOnEvent($g_iSettingsButtonCheckNow, "_OnEvent_SettingsCheckNow")

	$iY += 30

	; Одно поле под combo (выровнено по нему): синяя ссылка «Доступна версия X.XX» либо серая дата
	$g_iSettingsLabelUpdateInfo = GUICtrlCreateLabel("", 96, $iY, 190, 18, $SS_NOTIFY)
	GUICtrlSetOnEvent($g_iSettingsLabelUpdateInfo, "_OnEvent_SettingsUpdateLink")
	; Subclass для курсора-руки при наведении на ссылку
	If $g_hLinkSubclassCB = 0 Then $g_hLinkSubclassCB = DllCallbackRegister("_LinkSubclassProc", "lresult", "hwnd;uint;wparam;lparam;uint_ptr;dword_ptr")
	_WinAPI_SetWindowSubclass(GUICtrlGetHandle($g_iSettingsLabelUpdateInfo), DllCallbackGetPtr($g_hLinkSubclassCB), 1, 0)

	$iY += 34

	; Сброс кеша и Закрыть в одну строку. Описание кэша — в подсказке кнопки.
	$g_iSettingsButtonClearCache = GUICtrlCreateButton(Lang("GUI", "ClearCache", "Clear cache"), 16, $iY, 120, 26)
	GUICtrlSetTip($g_iSettingsButtonClearCache, Lang("Cache", "Desc", "Stores video resolutions and detected sync offsets. Clear it if cached data looks stale."))
	Local $iButtonClose = GUICtrlCreateButton(Lang("GUI", "OK", "OK"), 196, $iY, 90, 26)

	GUICtrlSetOnEvent($iButtonClose,                "_OnEvent_SettingsOk")
	GUICtrlSetOnEvent($g_iSettingsButtonClearCache, "_OnEvent_SettingsClearCache")
	GUISetOnEvent($GUI_EVENT_CLOSE, "_OnEvent_SettingsClose", $g_hSettingsGui)

	_ApplyThemeToSettingsGui()
	_SettingsUpdateInfoText() ; заполнить поле (ссылка/дата) с учётом темы

	; Иконка кнопки проверки (круг со стрелками), перекрашена под тему
	_DestroyImgList($g_hSettingsCheckImgList)
	$g_hSettingsCheckImgList = _SetButtonCenterIcon($g_iSettingsButtonCheckNow, "Refresh.png", 16)

	GUISetState(@SW_SHOW, $g_hSettingsGui)
EndFunc   ;==>_SettingsWindow


Func _ApplyThemeToSettingsGui()
	If $g_hSettingsGui = 0 Then Return

	; UDF 3.0.0 держит ОДИН глобальный реестр subclass-кнопок ($__DM_g_aButtonSub) на весь
	; процесс. На Win10 (не 24H2+) радио/чекбоксы главного окна subclass'атся, и их ButtonProc
	; читает $__DM_g_aButtonSub[$pData] без проверки границ. _GUIDarkTheme_GUISetDarkTheme при
	; КАЖДОМ вызове обнуляет этот массив ($__DM_g_aButtonSub[1][3], $__DM_g_iButtonCount = 0).
	; Тема окна настроек затирает его, у настроек subclass-кнопок нет — массив остаётся [1][3],
	; а радио главного окна хранят $pData = 1..4. Любая их перерисовка → индекс вне границ →
	; фатальный краш AutoIt. Сохраняем/восстанавливаем реестр вокруг покраски настроек.
	Local $aButtonSubSaved = $__DM_g_aButtonSub
	Local $iButtonCountSaved = $__DM_g_iButtonCount

	_GUIDarkTheme_GUISetDarkTheme($g_hSettingsGui, $g_bDarkMode)
	_GUIDarkTheme_GUICtrlAllSetDarkTheme($g_hSettingsGui, $g_bDarkMode)

	$__DM_g_aButtonSub = $aButtonSubSaved
	$__DM_g_iButtonCount = $iButtonCountSaved

	GUISetBkColor($g_iClrBg, $g_hSettingsGui)

	; Лейблы — прозрачный фон + цвет текста по палитре
	GUICtrlSetBkColor($g_iSettingsLabelLang,  $GUI_BKCOLOR_TRANSPARENT)
	GUICtrlSetBkColor($g_iSettingsLabelTheme, $GUI_BKCOLOR_TRANSPARENT)
	GUICtrlSetBkColor($g_iSettingsLabelUpdates,    $GUI_BKCOLOR_TRANSPARENT)
	GUICtrlSetBkColor($g_iSettingsLabelUpdateInfo, $GUI_BKCOLOR_TRANSPARENT)
	GUICtrlSetColor($g_iSettingsLabelLang,  $g_iClrFg)
	GUICtrlSetColor($g_iSettingsLabelTheme, $g_iClrFg)
	GUICtrlSetColor($g_iSettingsLabelUpdates,   $g_iClrFg)
	; Цвет поля обновлений (синий/серый) выставляет _SettingsUpdateInfoText
EndFunc   ;==>_ApplyThemeToSettingsGui


Func _OnEvent_SettingsClose()
	If $g_hSettingsGui = 0 Then Return
	_DestroyImgList($g_hSettingsCheckImgList) ; imagelist иконки кнопки проверки
	If $g_hLinkSubclassCB Then _WinAPI_RemoveWindowSubclass(GUICtrlGetHandle($g_iSettingsLabelUpdateInfo), DllCallbackGetPtr($g_hLinkSubclassCB), 1)
	GUIDelete($g_hSettingsGui)
	$g_hSettingsGui = 0
	$g_iSettingsButtonClearCache = 0
EndFunc   ;==>_OnEvent_SettingsClose


Func _OnEvent_SettingsOk()
	If $g_hSettingsGui = 0 Then Return

	; Язык
	Local $sSelectedLang = GUICtrlRead($g_iSettingsComboLang)
	Local $sLangCode = ($sSelectedLang = "Русский") ? "Russian" : "English"
	Local $bLangChanged = ($sLangCode <> $g_sCurrentLang)
	If $bLangChanged Then IniWrite($gc_sPathIni, "Settings", "Language", $sLangCode)

	; Тема
	Local $sSelectedTheme = GUICtrlRead($g_iSettingsComboTheme)
	Local $sTheme = "System"
	If $sSelectedTheme = Lang("GUI", "ThemeLight", "Light") Then $sTheme = "Light"
	If $sSelectedTheme = Lang("GUI", "ThemeDark", "Dark") Then $sTheme = "Dark"
	Local $bThemeChanged = ($sTheme <> $g_sTheme)
	If $bThemeChanged Then
		$g_sTheme = $sTheme
		IniWrite($gc_sPathIni, "Settings", "Theme", $sTheme)
	EndIf

	; Периодичность проверки обновлений (читаем до удаления окна)
	Local $sSelectedFreq = GUICtrlRead($g_iSettingsComboUpdates)
	Local $sFreqCode = "weekly"
	If $sSelectedFreq = Lang("Updates", "FreqNever", "Never") Then $sFreqCode = "never"
	If $sSelectedFreq = Lang("Updates", "FreqAlways", "At startup") Then $sFreqCode = "always"
	If $sSelectedFreq = Lang("Updates", "FreqMonthly", "Once a month") Then $sFreqCode = "monthly"
	IniWrite($gc_sPathIni, "Settings", "UpdateCheckFreq", $sFreqCode)

	_OnEvent_SettingsClose()

	If $bLangChanged Then
		_InitLanguage()
		_ApplyLanguage() ; ставит тексты + внутри зовёт _ApplyTheme и _RenderHeader
	EndIf
	If $bThemeChanged Then
		_InitTheme()
		If Not $bLangChanged Then
			_ApplyTheme()
			_RenderHeader()
		EndIf
	EndIf
	If $bLangChanged Or $bThemeChanged Then _
		DllCall("user32.dll", "bool", "RedrawWindow", "hwnd", $g_hGui, "ptr", 0, "handle", 0, "uint", $RDW_INVALIDATE + $RDW_UPDATENOW + $RDW_ALLCHILDREN)
EndFunc   ;==>_OnEvent_SettingsOk


Func _OnEvent_SettingsClearCache()
	Local $mEmpty[]
	$g_oCache = $mEmpty
	FileDelete($gc_sPathCache)
	_EnsureUtf16File($gc_sPathCache)
	GUICtrlSetState($g_iSettingsButtonClearCache, $GUI_DISABLE)

	; В режиме «Авто» смещение бралось из кеша — очищаем поле, статус и команду
	If GUICtrlRead($g_iRadioOffsetAuto) = $GUI_CHECKED Then
		GUICtrlSetData($g_iInputOffset, "")
		_SetSyncStatus("NOTRUN")
		_UpdateCommandField()
	EndIf
EndFunc   ;==>_OnEvent_SettingsClearCache


Func _OnEvent_SettingsUpdatesCombo()
	_SettingsUpdateInfoText()
EndFunc   ;==>_OnEvent_SettingsUpdatesCombo


; Заполняет поле группы «Проверка обновлений»: синяя ссылка при наличии обновления;
; иначе серая дата последней проверки; при «Никогда» (без обновления) поле скрыто.
Func _SettingsUpdateInfoText()
	If $g_hSettingsGui = 0 Then Return

	If $g_sAvailableVersion <> "" Then
		GUICtrlSetData($g_iSettingsLabelUpdateInfo, StringFormat(Lang("Updates", "AvailableVersion", "Version %s is available"), $g_sAvailableVersion))
		GUICtrlSetColor($g_iSettingsLabelUpdateInfo, $gc_iClrAccent) ; кликабельно
		GUICtrlSetState($g_iSettingsLabelUpdateInfo, $GUI_SHOW)
	ElseIf GUICtrlRead($g_iSettingsComboUpdates) = Lang("Updates", "FreqNever", "Never") Then
		GUICtrlSetState($g_iSettingsLabelUpdateInfo, $GUI_HIDE) ; проверки выключены — дата нерелевантна
	Else
		Local $sLast = IniRead($gc_sPathIni, "Settings", "LastUpdateCheck", "")
		If $sLast = "" Then $sLast = Lang("Updates", "LastCheckNever", "never") ; в ini уже YYYY.MM.DD
		GUICtrlSetData($g_iSettingsLabelUpdateInfo, Lang("Updates", "LastCheck", "Last check:") & " " & $sLast)
		GUICtrlSetColor($g_iSettingsLabelUpdateInfo, $g_iClrInfo) ; неактивный серый
		GUICtrlSetState($g_iSettingsLabelUpdateInfo, $GUI_SHOW)
	EndIf
EndFunc   ;==>_SettingsUpdateInfoText


; Сравнивает версии вида "1.07" покомпонентно как числа.
; Возвращает -1 ($s1 < $s2), 0 (равны), 1 ($s1 > $s2). Недостающие компоненты считаются нулём.
Func _CompareVersions($s1, $s2)
	Local $a1 = StringSplit($s1, ".", 2)
	Local $a2 = StringSplit($s2, ".", 2)
	Local $iMax = (UBound($a1) > UBound($a2)) ? UBound($a1) : UBound($a2)
	For $i = 0 To $iMax - 1
		Local $iN1 = ($i < UBound($a1)) ? Number($a1[$i]) : 0
		Local $iN2 = ($i < UBound($a2)) ? Number($a2[$i]) : 0
		If $iN1 < $iN2 Then Return -1
		If $iN1 > $iN2 Then Return 1
	Next
	Return 0
EndFunc   ;==>_CompareVersions


; Качает raw Include\AppConstants.au3 из ветки main и достаёт значение $gc_sAppVersion.
; Возвращает строку версии или "" с @error при сбое сети/парсинга.
Func _FetchRemoteVersion()
	Local $bData = InetRead($gc_sVerCheckUrl, $INET_FORCERELOAD)
	If @error Or BinaryLen($bData) = 0 Then Return SetError(1, 0, "")

	Local $sText = BinaryToString($bData, $SB_UTF8)
	Local $aMatch = StringRegExp($sText, '\$gc_sAppVersion\s*=\s*"([^"]+)"', 3)
	If Not IsArray($aMatch) Then Return SetError(2, 0, "")

	Return $aMatch[0]
EndFunc   ;==>_FetchRemoteVersion


; Проверяет обновления согласно периодичности из ini. При наличии более новой
; версии выставляет $g_sAvailableVersion и обновляет подвал. Без всплывающих окон.
Func _CheckUpdates()
	Local $sFreq = IniRead($gc_sPathIni, "Settings", "UpdateCheckFreq", "weekly")
	If $sFreq = "never" Then Return

	Local $sLast = IniRead($gc_sPathIni, "Settings", "LastUpdateCheck", "")
	Local $bShouldCheck = True
	If $sLast <> "" Then
		Local $sLastSlash = StringReplace($sLast, ".", "/") ; _DateDiff требует YYYY/MM/DD
		Switch $sFreq
			Case "weekly"
				$bShouldCheck = (_DateDiff("D", $sLastSlash, _NowCalc()) >= 7)
			Case "monthly"
				$bShouldCheck = (_DateDiff("D", $sLastSlash, _NowCalc()) >= 30)
				; "always" - всегда True
		EndSwitch
	EndIf
	If Not $bShouldCheck Then Return

	Local $sRemote = _FetchRemoteVersion()
	IniWrite($gc_sPathIni, "Settings", "LastUpdateCheck", StringReplace(_NowCalcDate(), "/", ".")) ; дата проверки даже при сбое сети
	If @error Or $sRemote = "" Then Return

	If _CompareVersions($gc_sAppVersion, $sRemote) < 0 Then
		$g_sAvailableVersion = $sRemote
		_ApplyUpdateNotice()
		_StartSettingsPulse() ; мигаем шестерёнкой — обновление найдено при старте
	EndIf
EndFunc   ;==>_CheckUpdates


; При наличии обновления дописывает в заголовок окна пометку «доступна X.XX».
Func _ApplyUpdateNotice()
	If $g_sAvailableVersion = "" Then Return
	WinSetTitle($g_hGui, "", $gc_sAppName & " " & Lang("Updates", "TitleUpdate", "[Update available]"))
EndFunc   ;==>_ApplyUpdateNotice


; Запускает мигание кнопки настроек (привлечь внимание к доступному обновлению).
Func _StartSettingsPulse()
	If $g_bPulseActive Then Return
	$g_bPulseActive = True
	$g_bPulsePhase = True
	AdlibRegister("_PulseSettingsButton", 600)
EndFunc   ;==>_StartSettingsPulse


; Останавливает мигание и возвращает обычную иконку шестерёнки.
Func _StopSettingsPulse()
	If Not $g_bPulseActive Then Return
	AdlibUnRegister("_PulseSettingsButton")
	$g_bPulseActive = False
	_DestroyImgList($g_hSettingsImgList)
	$g_hSettingsImgList = _SetButtonCenterIcon($g_iButtonSettings, "Settings.png", 18)
EndFunc   ;==>_StopSettingsPulse


; Тик мигания: переключает иконку шестерёнки между цветом темы и акцентным синим.
Func _PulseSettingsButton()
	_DestroyImgList($g_hSettingsImgList)
	$g_hSettingsImgList = _SetButtonCenterIcon($g_iButtonSettings, "Settings.png", 18, $g_bPulsePhase ? $gc_iClrAccent : _IconTint())
	$g_bPulsePhase = Not $g_bPulsePhase
EndFunc   ;==>_PulseSettingsButton


; Клик по ссылке обновления в настройках открывает страницу релизов.
Func _OnEvent_SettingsUpdateLink()
	If $g_sAvailableVersion <> "" Then ShellExecute($gc_sReleasesUrl)
EndFunc   ;==>_OnEvent_SettingsUpdateLink


; Subclass поля обновлений: курсор-рука при наведении, когда поле — кликабельная ссылка.
Func _LinkSubclassProc($hWnd, $iMsg, $wParam, $lParam, $iId, $dwData)
	#forceref $iId, $dwData
	If $iMsg = 0x0020 And $g_sAvailableVersion <> "" Then ; WM_SETCURSOR, только для ссылки
		_WinAPI_SetCursor(_WinAPI_LoadCursor(0, 32649)) ; IDC_HAND
		Return 1
	EndIf
	Return _WinAPI_DefSubclassProc($hWnd, $iMsg, $wParam, $lParam)
EndFunc   ;==>_LinkSubclassProc


; Кнопка «Проверить сейчас»: запускает проверку немедленно, минуя интервал.
Func _OnEvent_SettingsCheckNow()
	If $g_hSettingsGui = 0 Then Return

	; Кнопка иконочная — текст не трогаем, индикация через блокировку на время запроса
	GUICtrlSetState($g_iSettingsButtonCheckNow, $GUI_DISABLE)

	Local $sRemote = _FetchRemoteVersion()
	IniWrite($gc_sPathIni, "Settings", "LastUpdateCheck", StringReplace(_NowCalcDate(), "/", "."))
	If Not @error And $sRemote <> "" And _CompareVersions($gc_sAppVersion, $sRemote) < 0 Then
		$g_sAvailableVersion = $sRemote
		_ApplyUpdateNotice()
	EndIf

	GUICtrlSetState($g_iSettingsButtonCheckNow, $GUI_ENABLE)
	If $g_hSettingsGui <> 0 Then _SettingsUpdateInfoText() ; обновить поле (ссылка/дата)
EndFunc   ;==>_OnEvent_SettingsCheckNow


; === Ресурсы и завершение ===

; Освобождение GDI/HIMAGELIST handle'ов — срабатывает при любом пути выхода.
Func _Cleanup()
	If $g_bPulseActive Then AdlibUnRegister("_PulseSettingsButton")
	_HeaderDisposeBitmap($g_hLogoBitmap)
	_DestroyImgList($g_hCompareImgList)
	_DestroyImgList($g_hSettingsImgList)
	_DestroyImgList($g_hSwapImgList)
	_DestroyImgList($g_hSettingsCheckImgList)
	If $g_hLinkSubclassCB Then
		DllCallbackFree($g_hLinkSubclassCB)
		$g_hLinkSubclassCB = 0
	EndIf
EndFunc   ;==>_Cleanup


; Освобождает HIMAGELIST и обнуляет переменную (ByRef).
Func _DestroyImgList(ByRef $hList)
	If $hList Then
		DllCall("comctl32.dll", "int", "ImageList_Destroy", "handle", $hList)
		$hList = 0
	EndIf
EndFunc   ;==>_DestroyImgList
