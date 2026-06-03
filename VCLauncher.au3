#pragma compile(Out, VCLauncher.exe)
#pragma compile(Icon, Assets\Icon\Icon.ico)

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

#include "Include\GUIDarkTheme.au3"
#include "Include\FontHelper.au3"
#include "Include\HeaderHelper.au3"

Opt("GUIOnEventMode", 1)

; Константы приложения
Global Const $gc_sAppName = "VCLauncher 1.06"
Global Const $gc_sVcVersion = '20260502 "valencia"' ; версия video-compare, отображается в шапке
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
Global Const $gc_iClrBgDark = 0x202020
Global Const $gc_iClrBgLight = 0xF0F0F0
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
Global $g_iLabelOffset, $g_iRadioOffsetAuto, $g_iRadioOffsetManual, $g_iLabelSyncStatus, $g_iInputOffset

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
Global $g_iClrBg, $g_iClrFg, $g_iClrInfo, $g_iClrInput, $g_iClrSep
; Остальные переменные (чтение и нормализация путей; относительные пути считаем от папки скрипта)
Global $g_sVideoFile1 = _NormalizePath(IniRead($gc_sPathIni, "LastDirs", "Video1", ""))
Global $g_sVideoFile2 = _NormalizePath(IniRead($gc_sPathIni, "LastDirs", "Video2", ""))


_ResolveToolPaths()
_InitLanguage()
_InitTheme()

_CheckToolExists($g_sPathVideoCompare, "video-compare.exe")
_CheckToolExists($g_sPathSync, "Sync.exe")

OnAutoItExitRegister("_Cleanup")

_MainGUI()
_DefineEvents()

While 1
	Sleep(50)
WEnd


; Освобождение GDI/HIMAGELIST handle'ов — срабатывает при любом пути выхода.
Func _Cleanup()
	_HeaderDisposeBitmap($g_hLogoBitmap)
	_DestroyImgList($g_hCompareImgList)
	_DestroyImgList($g_hSettingsImgList)
	_DestroyImgList($g_hSwapImgList)
EndFunc   ;==>_Cleanup


; Освобождает HIMAGELIST и обнуляет переменную (ByRef).
Func _DestroyImgList(ByRef $hList)
	If $hList Then
		DllCall("comctl32.dll", "int", "ImageList_Destroy", "handle", $hList)
		$hList = 0
	EndIf
EndFunc   ;==>_DestroyImgList


Func _MainGUI()
	$g_hGui = GUICreate($gc_sAppName, $gc_iGuiWidth, $gc_iGuiHeight, -1, -1, $WS_SIZEBOX + $WS_SYSMENU + $WS_MINIMIZEBOX, $WS_EX_ACCEPTFILES)

	_ApplyAppFont()

	; Шапка: логотип слева + шпаргалка горячих клавиш справа (рисуется в _RenderHeader).
	; Pic ýже окна — справа остаётся полоса фона GUI под кнопку настроек.
	$g_iLogoPic = GUICtrlCreatePic("", 0, 0, $gc_iHeaderW, $gc_iHeaderH)

	; Кнопка настроек лежит на фоне GUI правее картинки-шапки, не перекрывая Pic.
	$g_iButtonSettings = GUICtrlCreateButton("", $gc_iGuiWidth - 40, 8, 30, 30) ; иконка-шестерёнка ставится ниже (PNG)
	GUICtrlSetTip($g_iButtonSettings, Lang("GUI", "Settings", "Settings"))

	; Разделитель сразу под шапкой (как перед блоком «Смещение»)
	$g_iSeparatorTop = GUICtrlCreateLabel("", 0, $gc_iHeaderH, $gc_iGuiWidth, 1)
	GUICtrlSetBkColor($g_iSeparatorTop, 0xC0C0C0)

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
	GUICtrlSetColor($g_iLabelInfo1, 0x808080)

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
	GUICtrlSetColor($g_iLabelInfo2, 0x808080)

	; --- Поменять местами ---
	$g_iButtonSwap = GUICtrlCreateButton("", $gc_iGuiWidth - 60, $iY + 24, 26, 26) ; иконка-свап ставится ниже (PNG)
	GUICtrlSetTip($g_iButtonSwap, Lang("GUI", "SwapTip", "Swap files"))

	; --- Разделитель ---
	$g_iSeparator = GUICtrlCreateLabel("", 0, $iY + 100, $gc_iGuiWidth, 1)
	GUICtrlSetBkColor($g_iSeparator, 0xC0C0C0)

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
	$g_iLabelSyncStatus = GUICtrlCreateLabel("", 142, $iY + 115, $gc_iGuiWidth - 376, 20, $SS_NOTIFY)
	GUICtrlSetFont($g_iLabelSyncStatus, $g_iAppFontSize, 400, 0, $g_sAppFont)
	$g_iInputOffset = GUICtrlCreateInput("", $gc_iGuiWidth - 150, $iY + 113, 60, 20)
	GUICtrlSetState($g_iInputOffset, $GUI_DISABLE)

	; --- Команда ---
	$g_iLabelCommand = GUICtrlCreateLabel(Lang("GUI", "TabCommand", "Command"), 10, $iY + 172, 70, 20)
	$g_iEditCommand = GUICtrlCreateEdit("", 90, $iY + 171, $gc_iGuiWidth - 180, 70, BitOR($ES_MULTILINE, $ES_AUTOVSCROLL, $WS_VSCROLL))

	; --- Кнопка «Сравнить» ---
	Local Const $BS_MULTILINE = 0x2000
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

	; Подсвечиваем соответствующий элемент
	If $iCtrlID = $g_iInput1 Then
		GUICtrlSetBkColor($g_iInput1, $COLOR_SKYBLUE)
	ElseIf $iCtrlID = $g_iInput2 Then
		GUICtrlSetBkColor($g_iInput2, $COLOR_SKYBLUE)
	EndIf

	; Через небольшую задержку восстанавливаем исходный вид после drop события
	AdlibRegister("_RestoreControlsStyle", 100)

	Return $GUI_RUNDEFMSG
EndFunc   ;==>_OnEvent_WM_DROPFILES


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

		; Читаем сдвиг из поля ввода
		Local $sOffset = GUICtrlRead($g_iInputOffset)
		Local $iOffsetMs = ($sOffset <> "") ? Int($sOffset) : 0

		; Обновляем поле команды
		Local $bIsVertical = (GUICtrlRead($g_iRadioVertical) = $GUI_CHECKED)
		Local $sCmdLine = _BuildVideoCompareCommand($aInfo1, $aInfo2, $aCropArgs, $bIsVertical, $iOffsetMs)
		GUICtrlSetData($g_iEditCommand, $sCmdLine)
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
	If Not FileExists($g_sVideoFile1) Or Not FileExists($g_sVideoFile2) Then
		GUICtrlSetData($g_iEditCommand, "")
		Return
	EndIf
	Local $aVideo1Info = _GetVideoInfo($g_sVideoFile1)
	Local $aVideo2Info = _GetVideoInfo($g_sVideoFile2)
	; Без валидного разрешения команду не строим — иначе деление на ноль в _CalculateCropArgs
	If $aVideo1Info[0] <= 0 Or $aVideo2Info[0] <= 0 Then
		GUICtrlSetData($g_iEditCommand, "")
		Return
	EndIf
	Local $aCropArgs = _CalculateCropArgs($aVideo1Info, $aVideo2Info)
	Local $sOffset = GUICtrlRead($g_iInputOffset)
	Local $iOffsetMs = ($sOffset <> "") ? Int($sOffset) : 0
	Local $bIsVertical = (GUICtrlRead($g_iRadioVertical) = $GUI_CHECKED)
	Local $sCmdLine = _BuildVideoCompareCommand($aVideo1Info, $aVideo2Info, $aCropArgs, $bIsVertical, $iOffsetMs)
	GUICtrlSetData($g_iEditCommand, $sCmdLine)
EndFunc   ;==>_UpdateCommandField


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

	; Кеш в памяти (прямая пара)
	If MapExists($g_oCache, $sMemKey) Then
		Local $vVal = $g_oCache[$sMemKey]
		$aResult[0] = True
		If _IsSyncMarker($vVal) Then
			$aResult[1] = 0
			$aResult[2] = $vVal
		Else
			$aResult[1] = Int($vVal)
			$aResult[2] = "OK"
		EndIf
		Return $aResult
	EndIf

	; Кеш в памяти (обратная пара — инвертируем сдвиг для OK, маркер копируем)
	Local $sMemKeyRev = _BuildPairKey($sFile2, $sFile1)
	If MapExists($g_oCache, $sMemKeyRev) Then
		Local $vVal = $g_oCache[$sMemKeyRev]
		$aResult[0] = True
		If _IsSyncMarker($vVal) Then
			$aResult[1] = 0
			$aResult[2] = $vVal
			$g_oCache[$sMemKey] = $vVal
		Else
			$aResult[1] = -Int($vVal)
			$aResult[2] = "OK"
			$g_oCache[$sMemKey] = $aResult[1]
		EndIf
		Return $aResult
	EndIf

	; Кеш на диске: находим индексы из [Info] и ищем в [Sync]
	Local $aInfo1 = _GetCacheInfo($sFile1)
	Local $aInfo2 = _GetCacheInfo($sFile2)
	If $aInfo1[0] = "" Or $aInfo2[0] = "" Then Return $aResult

	; Проверяем прямую пару
	Local $sSyncVal = IniRead($gc_sPathCache, "Sync", $aInfo1[0] & "|" & $aInfo2[0], "")
	If $sSyncVal <> "" Then
		$aResult[0] = True
		If _IsSyncMarker($sSyncVal) Then
			$aResult[1] = 0
			$aResult[2] = $sSyncVal
			$g_oCache[$sMemKey] = $sSyncVal
		Else
			$aResult[1] = Int($sSyncVal)
			$aResult[2] = "OK"
			$g_oCache[$sMemKey] = $aResult[1]
		EndIf
		Return $aResult
	EndIf

	; Обратная пара
	Local $sSyncValRev = IniRead($gc_sPathCache, "Sync", $aInfo2[0] & "|" & $aInfo1[0], "")
	If $sSyncValRev <> "" Then
		$aResult[0] = True
		If _IsSyncMarker($sSyncValRev) Then
			$aResult[1] = 0
			$aResult[2] = $sSyncValRev
			$g_oCache[$sMemKey] = $sSyncValRev
		Else
			$aResult[1] = -Int($sSyncValRev)
			$aResult[2] = "OK"
			$g_oCache[$sMemKey] = $aResult[1]
		EndIf
		Return $aResult
	EndIf

	Return $aResult
EndFunc   ;==>_LookupSyncCache


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
	While 1
		Local $sLine = StdoutRead($iPid)
		If @error Then ExitLoop
		$sStdout &= $sLine
	WEnd
	While 1
		Local $sLine = StderrRead($iPid)
		If @error Then ExitLoop
		$sStderr &= $sLine
	WEnd

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
	If $g_sLastSyncStatus <> "NOMATCH" And $g_sLastSyncStatus <> "TIMEOUT" And $g_sLastSyncStatus <> "ERROR" Then Return

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

	Local $aPipe = DllCall("kernel32.dll", "bool", "CreatePipe", "handle*", 0, "handle*", 0, "struct*", $tSA, "dword", 0)
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
		Local $aPeek = DllCall("kernel32.dll", "bool", "PeekNamedPipe", "handle", $hRead, "ptr", 0, "dword", 0, "ptr", 0, "struct*", $tAvail, "ptr", 0)
		If @error Or Not $aPeek[0] Then ExitLoop
		Local $iAvail = DllStructGetData($tAvail, 1)
		If $iAvail > 0 Then
			Local $aRF = DllCall("kernel32.dll", "bool", "ReadFile", "handle", $hRead, "struct*", $tBuf, "dword", 4096, "struct*", $tRead, "ptr", 0)
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
				Local $aPeek2 = DllCall("kernel32.dll", "bool", "PeekNamedPipe", "handle", $hRead, "ptr", 0, "dword", 0, "ptr", 0, "struct*", $tAvail, "ptr", 0)
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
	Local $sHeaderIcon = @ScriptDir & "\Assets\Icon\HeaderIcon.png"

	; Заголовки шапки под иконкой: VCLauncher, затем Video-compare и его версия отдельной строкой.
	Local $sTitleApp = $gc_sAppName
	Local $sTitleVc = "Video-compare"
	Local $sTitleVcVer = $gc_sVcVersion
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
	Return $g_bDarkMode ? 0xE6E6E6 : 0x303030
EndFunc   ;==>_IconTint


; Ставит PNG-иконку из Assets\Icon по центру кнопки, перекрашенную под тему. Возвращает HIMAGELIST.
Func _SetButtonCenterIcon($iCtrl, $sPngName, $iIconSize)
	Local $sPath = @ScriptDir & "\Assets\Icon\" & $sPngName
	Local $hIcon = _LoadHIconFromPng($sPath, $iIconSize, _IconTint())
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
	Local $sPath = @ScriptDir & "\Assets\CompareIcons\" & $sIconName
	Local Const $iIconSize = 32

	Local $hIcon = _LoadHIconFromPng($sPath, $iIconSize, _IconTint()) ; перекраска под тему
	If Not $hIcon Then Return

	Local $hNew = _ApplyHIconToButton($g_iButtonCompare, $hIcon, $iIconSize)
	If Not $hNew Then Return

	; Освобождаем предыдущий imagelist после замены
	If $g_hCompareImgList Then
		DllCall("comctl32.dll", "int", "ImageList_Destroy", "handle", $g_hCompareImgList)
	EndIf
	$g_hCompareImgList = $hNew
EndFunc   ;==>_UpdateCompareButtonIcon


Func _RestoreControlsStyle()
	AdlibUnRegister("_RestoreControlsStyle")

	; Восстанавливаем цвет фона input полей
	GUICtrlSetBkColor($g_iInput1, $g_iClrInput)
	GUICtrlSetBkColor($g_iInput2, $g_iClrInput)

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
			"SyncMethod=audio" & @CRLF
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
		$g_iClrInput = 0x3C3C3C
		$g_iClrSep = $COLOR_BORDER
	Else
		$g_iClrBg = $gc_iClrBgLight
		$g_iClrFg = 0x000000
		$g_iClrInfo = 0x808080
		$g_iClrInput = 0xFFFFFF
		$g_iClrSep = 0xC0C0C0
	EndIf
EndFunc   ;==>_SetPalette


Func _ApplyTheme()
	_ResolveDarkMode()

	; Сравниваем с реально применённой к GUI темой, а не с $g_bDarkMode до вызова:
	; _InitTheme() может быть вызван выше по стеку и уже обновить $g_bDarkMode.
	Local $bNeedSwitch = (Not $g_bThemeInitialized) Or ($g_bAppliedDark <> $g_bDarkMode)

	; Применяем UDF-тему ко всему GUI. Не используем _GUIDarkTheme_SwitchTheme —
	; он определяет направление по системной теме, а не по нашему выбору.
	; Важно: _GUIDarkTheme_ApplyLight меняет внутренние UDF-глобалы на светлые значения,
	; но _GUIDarkTheme_ApplyDark их обратно не восстанавливает. Поэтому перед переключением
	; в тёмную тему явно возвращаем тёмную палитру UDF (те же значения, что задаёт _SwitchTheme).
	If $bNeedSwitch Then
		If $g_bThemeInitialized Then __GUIDarkTheme_SubclassCleanup()
		If $g_bDarkMode Then
			$g_iBkColor = 0x1C1C1C
			$COLOR_BG_DARK = 0x121212
			$COLOR_TEXT_LIGHT = 0xE0E0E0
			$COLOR_CONTROL_BG = 0x202020
			$COLOR_BORDER_LIGHT = 0xB0B0B0
			$COLOR_BORDER = 0x3F3F3F
			_GUIDarkTheme_ApplyDark($g_hGui, True)
		Else
			_GUIDarkTheme_ApplyLight($g_hGui)
		EndIf
		$g_bThemeInitialized = True
		$g_bAppliedDark = $g_bDarkMode
	EndIf

	; Палитра для кастомных элементов
	_SetPalette()

	; UDF ставит фон окна 0x121212 — в тёмной теме делаем его светлее
	If $g_bDarkMode Then GUISetBkColor($g_iClrBg, $g_hGui)

	; Info-лейблы — серый оттенок
	GUICtrlSetColor($g_iLabelInfo1, $g_iClrInfo)
	GUICtrlSetColor($g_iLabelInfo2, $g_iClrInfo)

	; Поле ввода сдвига
	GUICtrlSetColor($g_iInputOffset, $g_iClrFg)
	GUICtrlSetBkColor($g_iInputOffset, $g_iClrInput)

	; Статус-иконка sync — прозрачный фон. Цвет символа задаётся в _SetSyncStatus.
	GUICtrlSetBkColor($g_iLabelSyncStatus, $GUI_BKCOLOR_TRANSPARENT)
	; При смене темы пересчитываем цвет под текущий статус
	If $g_sLastSyncStatus <> "" Then _SetSyncStatus($g_sLastSyncStatus, $g_iLastSyncOffset)

	; Иконки кнопок перекрашиваем под цвет текста темы
	_ApplyButtonIcons()

	; Горизонтальный разделитель
	GUICtrlSetBkColor($g_iSeparator, $g_iClrSep)
	GUICtrlSetBkColor($g_iSeparatorTop, $g_iClrSep)

	; Иконки подсказок должны соответствовать выбранной теме (Dark/Light)
	If $bNeedSwitch Then _RenderHeader()

	; Финальная полная перерисовка окна
	; RDW_INVALIDATE=0x1, RDW_UPDATENOW=0x100, RDW_ALLCHILDREN=0x80
	DllCall("user32.dll", "bool", "RedrawWindow", "hwnd", $g_hGui, "ptr", 0, "handle", 0, "uint", $RDW_INVALIDATE + $RDW_UPDATENOW + $RDW_ALLCHILDREN)
EndFunc   ;==>_ApplyTheme




Func _SettingsWindow()
	; По центру основного окна (а не экрана)
	Local Const $iSetW = 320, $iSetH = 145
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

	$iY += 48

	; Сброс кеша и Закрыть в одну строку
	$g_iSettingsButtonClearCache = GUICtrlCreateButton(Lang("GUI", "ClearCache", "Clear cache"), 16, $iY, 120, 26)
	Local $iButtonClose          = GUICtrlCreateButton(Lang("GUI", "OK", "OK"), 214, $iY, 90, 26)

	GUICtrlSetOnEvent($iButtonClose,                "_OnEvent_SettingsOk")
	GUICtrlSetOnEvent($g_iSettingsButtonClearCache, "_OnEvent_SettingsClearCache")
	GUISetOnEvent($GUI_EVENT_CLOSE, "_OnEvent_SettingsClose", $g_hSettingsGui)

	_ApplyThemeToSettingsGui()

	GUISetState(@SW_SHOW, $g_hSettingsGui)
EndFunc   ;==>_SettingsWindow


Func _ApplyThemeToSettingsGui()
	If $g_hSettingsGui = 0 Then Return

	; ВАЖНО: _GUIDarkTheme_GUISetDarkTheme внутри перезаписывает глобал $g_hGui UDF'а
	; на переданный $hWnd. Сохраняем и восстанавливаем наш главный handle.
	Local $hMainSaved = $g_hGui
	_GUIDarkTheme_GUISetDarkTheme($g_hSettingsGui, $g_bDarkMode)
	_GUIDarkTheme_GUICtrlAllSetDarkTheme($g_hSettingsGui, $g_bDarkMode, True)
	$g_hGui = $hMainSaved

	GUISetBkColor($g_iClrBg, $g_hSettingsGui)

	; Лейблы — прозрачный фон + цвет текста по палитре
	GUICtrlSetBkColor($g_iSettingsLabelLang,  $GUI_BKCOLOR_TRANSPARENT)
	GUICtrlSetBkColor($g_iSettingsLabelTheme, $GUI_BKCOLOR_TRANSPARENT)
	GUICtrlSetColor($g_iSettingsLabelLang,  $g_iClrFg)
	GUICtrlSetColor($g_iSettingsLabelTheme, $g_iClrFg)
EndFunc   ;==>_ApplyThemeToSettingsGui


Func _OnEvent_SettingsClose()
	If $g_hSettingsGui = 0 Then Return
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


Func _MainGUI_Refresh()
	_ApplyTheme()
	_RenderHeader()
	; Перерисовка окна
	DllCall("user32.dll", "bool", "RedrawWindow", "hwnd", $g_hGui, "ptr", 0, "handle", 0, "uint", $RDW_INVALIDATE + $RDW_UPDATENOW + $RDW_ALLCHILDREN)
EndFunc   ;==>_MainGUI_Refresh
