#pragma compile(Out, VCLauncher.exe)
#pragma compile(Icon, Assets\icon.ico)

#NoTrayIcon
#RequireAdmin

#include <GUIConstantsEx.au3>
#include <WindowsConstants.au3>
#include <FileConstants.au3>
#include <WinAPI.au3>
#include <Math.au3>
#include <ColorConstants.au3>
#include <EditConstants.au3>
#include <ComboConstants.au3>
#include <GDIPlus.au3>
#include "Include\GUIDarkTheme.au3"

Opt("GUIOnEventMode", 1)

; Константы приложения
Global Const $sAppName = "VCLauncher 1.04"
Global Const $sPathIni = @ScriptDir & '\VCLauncher.ini'
Global Const $sPathCache = @ScriptDir & '\VCLauncher.cache'
Global Const $aSupportedExtensions[] = ["mp4", "avi", "mkv", "mov", "wmv", "webm", "mpg", "mpeg"]
Global Const $iGuiWidth = 500
Global Const $iHeaderH = 80 ; высота шапки с логотипом
Global Const $iGuiHeight = $iHeaderH + 266

; Переменные путей к инструментам (загружаются из ini)
Global $sPathVideoCompare = @ScriptDir & '\video-compare.exe'
Global $sPathSync = @ScriptDir & '\sync\dist\sync.exe'

; Создание ini по умолчанию, если отсутствует
_EnsureIniDefaults()

; Ссылки на элементы GUI
Global $hGui, $iInput1, $iInput2, $iButtonChoose1, $iButtonChoose2, $iButtonSwap, $iButtonCompare, $iButtonClearCache, $iRadioDirect, $iRadioVertical
Global $iLogoPic, $hLogoBitmap = 0 ; картинка-логотип в шапке и её HBITMAP для освобождения
Global $g_sAppFont = "MS Shell Dlg 2", $g_iAppFontSize = 9
; Кеш разрешений видео в пределах сессии. Ключ = путь|размер|mtime — автоматически
; инвалидируется, если файл на диске заменён.
Global $g_oInfoCache = ObjCreate("Scripting.Dictionary")
; Кеш результата sync: ключ = пара файлов + size + mtime, значение = сдвиг в мс
Global $g_oSyncCache = ObjCreate("Scripting.Dictionary")
Global $iLabel1, $iLabel2, $iLabelInfo1, $iLabelInfo2, $iLabelCompare, $iLabelCommand
Global $iEditCommand, $iComboLang, $iLabelSettings, $iLabelLang, $iLabelTheme, $iComboTheme
Global $iSeparator, $iSeparator2
Global $iLabelOffset, $iCheckSync, $iProgressSync, $iInputOffset

Global $g_sLangFile = "", $g_sCurrentLang = "", $g_oLangDict = Null

; Тема оформления
Global $g_sTheme = "System" ; System | Light | Dark
Global $g_bDarkMode = False
Global $g_bThemeInitialized = False ; флаг: применяли ли уже UDF-тему
Global $g_iClrBg, $g_iClrFg, $g_iClrInfo, $g_iClrInput, $g_iClrSep
; Остальные переменные (чтение и нормализация путей; относительные пути считаем от папки скрипта)
Global $sVideoFile1 = _NormalizePath(IniRead($sPathIni, "LastDirs", "Video1", ""))
Global $sVideoFile2 = _NormalizePath(IniRead($sPathIni, "LastDirs", "Video2", ""))


_ResolveToolPaths()
_InitLanguage()
_InitTheme()

If Not FileExists($sPathVideoCompare) Then
	MsgBox(48, $sAppName, StringReplace(Lang("Errors", "ToolNotFound", '"%TOOL%" not found.'), "%TOOL%", "video-compare.exe") & @CR & @CR & _
			Lang("Errors", "SettingsWillOpen", "The settings file will now be opened."))
	ShellExecute($sPathIni)
	Exit
EndIf

If Not FileExists($sPathSync) Then
	MsgBox(48, $sAppName, StringReplace(Lang("Errors", "ToolNotFound", '"%TOOL%" not found.'), "%TOOL%", "sync.exe") & @CR & @CR & _
			Lang("Errors", "SettingsWillOpen", "The settings file will now be opened."))
	ShellExecute($sPathIni)
	Exit
EndIf

_MainGUI()
_DefineEvents()

While 1
	Sleep(50)
WEnd


Func _MainGUI()
	$hGui = GUICreate($sAppName, $iGuiWidth, $iGuiHeight, -1, -1, $WS_SIZEBOX + $WS_SYSMENU + $WS_MINIMIZEBOX, $WS_EX_ACCEPTFILES)

	_ApplyAppFont()

	; Шапка: логотип слева + 4 строки горячих клавиш справа (рисуется в _RenderHeader)
	$iLogoPic = GUICtrlCreatePic("", 0, 0, $iGuiWidth, $iHeaderH)

	Local $iY = $iHeaderH + 10 ; отступ от нижнего края шапки

	; --- Видео 1 ---
	$iLabel1 = GUICtrlCreateLabel(Lang("GUI", "File1", "File 1"), 10, $iY + 2, 74, 20)

	$iInput1 = GUICtrlCreateInput("", 90, $iY, $iGuiWidth - 180, 20)
	GUICtrlSetState($iInput1, $GUI_DROPACCEPTED)
	If $sVideoFile1 <> "" Then GUICtrlSetData($iInput1, _GetFileName($sVideoFile1))

	$iButtonChoose1 = GUICtrlCreateButton(Lang("GUI", "Choose", "Choose"), $iGuiWidth - 82, $iY - 1, 74, 23)

	$iLabelInfo1 = GUICtrlCreateLabel(Lang("GUI", "FileNotSelected", "File not selected"), 90, $iY + 22, $iGuiWidth - 180, 15)
	GUICtrlSetColor($iLabelInfo1, 0x808080)

	; --- Видео 2 ---
	$iLabel2 = GUICtrlCreateLabel(Lang("GUI", "File2", "File 2"), 10, $iY + 50, 74, 20)

	$iInput2 = GUICtrlCreateInput("", 90, $iY + 48, $iGuiWidth - 180, 20)
	GUICtrlSetState($iInput2, $GUI_DROPACCEPTED)
	If $sVideoFile2 <> "" Then GUICtrlSetData($iInput2, _GetFileName($sVideoFile2))

	$iButtonSwap = GUICtrlCreateButton(ChrW(0x21C5), $iGuiWidth - 82, $iY + 23, 74, 23)
	GUICtrlSetTip($iButtonSwap, Lang("GUI", "SwapTip", "Swap files"))
	GUICtrlSetFont($iButtonSwap, $g_iAppFontSize + 1, 700, 0, $g_sAppFont)

	$iButtonChoose2 = GUICtrlCreateButton(Lang("GUI", "Choose", "Choose"), $iGuiWidth - 82, $iY + 47, 74, 23)

	$iLabelInfo2 = GUICtrlCreateLabel(Lang("GUI", "FileNotSelected", "File not selected"), 90, $iY + 70, $iGuiWidth - 180, 15)
	GUICtrlSetColor($iLabelInfo2, 0x808080)

	; --- Разделитель ---
	$iSeparator = GUICtrlCreateLabel("", 0, $iY + 93, $iGuiWidth, 1)
	GUICtrlSetBkColor($iSeparator, 0xC0C0C0)

	; --- Режим сравнения ---
	$iLabelCompare = GUICtrlCreateLabel(Lang("GUI", "CompareMode", "Compare:"), 10, $iY + 105, 74, 20)

	$iRadioDirect = GUICtrlCreateRadio(Lang("GUI", "CompareDirect", "Direct"), 90, $iY + 103, 130, 20)
	$iRadioVertical = GUICtrlCreateRadio(Lang("GUI", "CompareVertical", "Vertical"), 230, $iY + 103, 160, 20)
	GUICtrlSetState($iRadioDirect, $GUI_CHECKED)

	; --- Синхронизация ---
	$iLabelOffset = GUICtrlCreateLabel(Lang("GUI", "Offset", "Offset:"), 10, $iY + 130, 74, 20)
	$iCheckSync = GUICtrlCreateCheckbox(Lang("GUI", "SyncVideo", "Synchronize video"), 90, $iY + 128, 210, 20)
	GUICtrlSetState($iCheckSync, $GUI_CHECKED)
	$iInputOffset = GUICtrlCreateInput("", $iGuiWidth - 150, $iY + 128, 60, 20)
	GUICtrlSetState($iInputOffset, $GUI_DISABLE)
	$iProgressSync = GUICtrlCreateProgress($iGuiWidth - 150, $iY + 128, 60, 20)
	GUICtrlSetState($iProgressSync, $GUI_HIDE)

	; --- Команда ---
	$iLabelCommand = GUICtrlCreateLabel(Lang("GUI", "TabCommand", "Command"), 10, $iY + 159, 74, 20)

	$iEditCommand = GUICtrlCreateEdit("", 90, $iY + 156, $iGuiWidth - 180, 34, BitOR($ES_MULTILINE, $ES_AUTOVSCROLL))

	; --- Кнопка «Сравнить» ---
	Local Const $BS_MULTILINE = 0x2000
	$iButtonCompare = GUICtrlCreateButton(Lang("GUI", "Compare", "Compare"), $iGuiWidth - 82, $iY + 106, 74, 74, $BS_MULTILINE)
	_SetButtonIcon($iButtonCompare, @SystemDir & "\imageres.dll", 18, 32)

	; --- Разделитель 2 ---
	$iSeparator2 = GUICtrlCreateLabel("", 0, $iY + 195, $iGuiWidth, 1)
	GUICtrlSetBkColor($iSeparator2, 0xC0C0C0)

	; --- Настройки: язык, тема и сброс кеша (в одну строку) ---
	$iLabelSettings = GUICtrlCreateLabel(Lang("GUI", "TabSettings", "Settings:"), 10, $iY + 206, 74, 20)
	$iLabelLang = GUICtrlCreateLabel(Lang("GUI", "Language", "Language:"), 90, $iY + 206, 40, 20)
	$iComboLang = GUICtrlCreateCombo("", 130, $iY + 204, 90, 200, $CBS_DROPDOWNLIST)
	GUICtrlSetData($iComboLang, "English|Русский", ($g_sCurrentLang = "Russian") ? "Русский" : "English")
	_SetComboItemHeight($iComboLang, 17)

	$iLabelTheme = GUICtrlCreateLabel(Lang("GUI", "Theme", "Theme:"), 230, $iY + 206, 40, 20)
	$iComboTheme = GUICtrlCreateCombo("", 270, $iY + 204, 90, 200, $CBS_DROPDOWNLIST)
	_PopulateComboTheme()
	_SetComboItemHeight($iComboTheme, 17)

	$iButtonClearCache = GUICtrlCreateButton(Lang("GUI", "ClearCache", "Clear cache"), $iGuiWidth - 82, $iY + 203, 74, 23)

	_SetCtrlResizing()
	_RenderHeader()
	_ApplyTheme()

	GUISetState(@SW_SHOW)

	; Отложенная загрузка: окно уже видно, теперь подтягиваем данные
	_TryFillCachedOffset()
	_UpdateFilesInfo()
EndFunc   ;==>_MainGUI


Func _DefineEvents()
	GUICtrlSetOnEvent($iButtonChoose1, "_OnEvent_ButtonChoose")
	GUICtrlSetOnEvent($iButtonChoose2, "_OnEvent_ButtonChoose")
	GUICtrlSetOnEvent($iButtonSwap, "_OnEvent_ButtonSwap")
	GUICtrlSetOnEvent($iButtonCompare, "_OnEvent_ButtonCompare")
	GUICtrlSetOnEvent($iButtonClearCache, "_OnEvent_ButtonClearCache")
	GUICtrlSetOnEvent($iCheckSync, "_OnEvent_CheckSync")
	GUICtrlSetOnEvent($iRadioDirect, "_OnEvent_RadioChanged")
	GUICtrlSetOnEvent($iRadioVertical, "_OnEvent_RadioChanged")
	GUICtrlSetOnEvent($iComboLang, "_OnEvent_ComboLang")
	GUICtrlSetOnEvent($iComboTheme, "_OnEvent_ComboTheme")
	GUISetOnEvent($GUI_EVENT_CLOSE, "_OnEvent_GUI_EVENT_CLOSE")
	GUISetOnEvent($GUI_EVENT_DROPPED, "_OnEvent_GUI_EVENT_DROPPED")

	GUIRegisterMsg($WM_GETMINMAXINFO, "_OnEvent_WM_GETMINMAXINFO")
	GUIRegisterMsg($WM_COMMAND, "_OnEvent_WM_COMMAND")
	GUIRegisterMsg(0x0233, "_OnEvent_WM_DROPFILES") ; WM_DROPFILES
EndFunc   ;==>_DefineEvents


Func _OnEvent_WM_GETMINMAXINFO($hWnd, $iMsg, $wParam, $lParam)
	#forceref $iMsg, $wParam
	If $hWnd = $hGui Then
		Local $tMINMAXINFO = DllStructCreate("int;int;" & _
				"int MaxSizeX; int MaxSizeY;" & _
				"int MaxPositionX;int MaxPositionY;" & _
				"int MinTrackSizeX; int MinTrackSizeY;" & _
				"int MaxTrackSizeX; int MaxTrackSizeY", _
				$lParam)
		DllStructSetData($tMINMAXINFO, "MinTrackSizeX", $iGuiWidth) ; минимальная ширина окна
		DllStructSetData($tMINMAXINFO, "MinTrackSizeY", $iGuiHeight + 14) ; минимальная высота окна
	EndIf
	Return $GUI_RUNDEFMSG
EndFunc   ;==>_OnEvent_WM_GETMINMAXINFO


Func _OnEvent_ButtonChoose()
	Local $iButtonID = @GUI_CtrlId
	Local $sTitle, $sIniKey, $iInputCtrl

	If $iButtonID = $iButtonChoose1 Then
		$sTitle = Lang("Dialogs", "SelectVideo1", "Select video 1")
		$sIniKey = "Video1"
		$iInputCtrl = $iInput1
		Local $sCurrentFile = $sVideoFile1
	Else
		$sTitle = Lang("Dialogs", "SelectVideo2", "Select video 2")
		$sIniKey = "Video2"
		$iInputCtrl = $iInput2
		Local $sCurrentFile = $sVideoFile2
	EndIf

	Local $sFile = FileOpenDialog($sTitle, _PathGetDir($sCurrentFile), _GetVideoExtensionsFilter(), 1, _GetFileName($sCurrentFile))
	If Not @error And FileExists($sFile) Then
		If $iButtonID = $iButtonChoose1 Then
			$sVideoFile1 = $sFile
		Else
			$sVideoFile2 = $sFile
		EndIf
		GUICtrlSetData($iInputCtrl, _GetFileName($sFile))
		IniWrite($sPathIni, "LastDirs", $sIniKey, $sFile)
		GUICtrlSetData($iInputOffset, "")
		_TryFillCachedOffset()
		_UpdateFilesInfo()
	EndIf
EndFunc   ;==>_OnEvent_ButtonChoose


Func _OnEvent_ButtonCompare()
	If Not FileExists($sVideoFile1) Or Not FileExists($sVideoFile2) Then Return

	; Индикация «идёт запуск»: блокируем кнопку и меняем её надпись
	GUICtrlSetState($iButtonCompare, $GUI_DISABLE)
	GUICtrlSetData($iButtonCompare, Lang("GUI", "Running", "Running..."))

	; Синхронизация, если включена
	If GUICtrlRead($iCheckSync) = $GUI_CHECKED Then
		GUICtrlSetState($iInputOffset, $GUI_HIDE)
		GUICtrlSetState($iProgressSync, $GUI_SHOW)
		GUICtrlSetData($iProgressSync, 0)
		Local $iOffsetMs = _GetSyncOffset($sVideoFile1, $sVideoFile2)
		GUICtrlSetData($iProgressSync, 100)
		Sleep(100)
		GUICtrlSetData($iInputOffset, $iOffsetMs)
		GUICtrlSetState($iProgressSync, $GUI_HIDE)
		GUICtrlSetState($iInputOffset, $GUI_SHOW)
	EndIf

	_UpdateCommandField()

	Local $sCmdLine = StringReplace(StringReplace(GUICtrlRead($iEditCommand), @CR, " "), @LF, " ")
	$sCmdLine = StringStripWS($sCmdLine, 3)
	If $sCmdLine <> "" Then _RunVideoCompare($sCmdLine)

	; Возвращаем кнопку в исходное состояние
	GUICtrlSetData($iButtonCompare, Lang("GUI", "Compare", "Compare"))
	_UpdateFilesInfo()
EndFunc   ;==>_OnEvent_ButtonCompare


Func _OnEvent_ButtonClearCache()
	; Очищаем кеш в памяти
	$g_oSyncCache.RemoveAll()
	$g_oInfoCache.RemoveAll()

	; Удаляем файл кеша на диске
	FileDelete($sPathCache)

	; Очищаем поле сдвига
	GUICtrlSetData($iInputOffset, "")
EndFunc   ;==>_OnEvent_ButtonClearCache


Func _OnEvent_ButtonSwap()
	Local $sTmp = $sVideoFile1
	$sVideoFile1 = $sVideoFile2
	$sVideoFile2 = $sTmp

	IniWrite($sPathIni, "LastDirs", "Video1", $sVideoFile1)
	IniWrite($sPathIni, "LastDirs", "Video2", $sVideoFile2)

	GUICtrlSetData($iInput1, $sVideoFile1 = "" ? "" : _GetFileName($sVideoFile1))
	GUICtrlSetData($iInput2, $sVideoFile2 = "" ? "" : _GetFileName($sVideoFile2))

	; Инвертируем сдвиг вместо повторного поиска
	Local $sOffset = GUICtrlRead($iInputOffset)
	If $sOffset <> "" Then
		GUICtrlSetData($iInputOffset, -Int($sOffset))
	EndIf
	_UpdateFilesInfo()
EndFunc   ;==>_OnEvent_ButtonSwap


Func _OnEvent_RadioChanged()
	_UpdateCommandField()
EndFunc   ;==>_OnEvent_RadioChanged


Func _OnEvent_CheckSync()
	If GUICtrlRead($iCheckSync) = $GUI_CHECKED Then
		GUICtrlSetState($iInputOffset, $GUI_DISABLE)
		_TryFillCachedOffset()
	Else
		GUICtrlSetState($iInputOffset, $GUI_ENABLE)
	EndIf
	_UpdateCommandField()
EndFunc   ;==>_OnEvent_CheckSync


Func _OnEvent_GUI_EVENT_CLOSE()
	_DisposeLogo()
	Exit
EndFunc   ;==>_OnEvent_GUI_EVENT_CLOSE


Func _OnEvent_WM_DROPFILES($hWnd, $iMsg, $wParam, $lParam)
	#forceref $hWnd, $iMsg, $wParam, $lParam

	; Определяем над каким элементом находится курсор
	Local $aCursorInfo = GUIGetCursorInfo($hGui)
	If @error Then Return $GUI_RUNDEFMSG

	Local $iCtrlID = $aCursorInfo[4]

	; Подсвечиваем соответствующий элемент
	If $iCtrlID = $iInput1 Then
		GUICtrlSetBkColor($iInput1, $COLOR_SKYBLUE)
	ElseIf $iCtrlID = $iInput2 Then
		GUICtrlSetBkColor($iInput2, $COLOR_SKYBLUE)
	EndIf

	; Через небольшую задержку восстанавливаем исходный вид после drop события
	AdlibRegister("_RestoreControlsStyle", 100)

	Return $GUI_RUNDEFMSG
EndFunc   ;==>_OnEvent_WM_DROPFILES


Func _OnEvent_GUI_EVENT_DROPPED()
	Local $iDropId = @GUI_DropId
	Local $sDropFile = @GUI_DragFile

	; Проверяем, что файл брошен на один из input полей
	If $iDropId <> $iInput1 And $iDropId <> $iInput2 Then Return

	; Проверяем существование файла
	If Not FileExists($sDropFile) Then Return

	; Проверяем расширение файла
	If Not _IsValidVideoExtension($sDropFile) Then
		Local $sExt = StringLower(StringRegExpReplace($sDropFile, '^.*\.', ''))
		MsgBox(48, $sAppName, Lang("Errors", "UnsupportedFormat", "Unsupported file format:") & " " & $sExt & @CR & @CR & _
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
	If $iID = $iInputOffset Then
		_UpdateCommandField()
		Return $GUI_RUNDEFMSG
	EndIf

	If $iID <> $iInput1 And $iID <> $iInput2 Then Return $GUI_RUNDEFMSG

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
		If $iID = $iInput1 And FileExists($sVideoFile1) Then
			$sContextDir = _PathGetDir($sVideoFile1)
		ElseIf $iID = $iInput2 And FileExists($sVideoFile2) Then
			$sContextDir = _PathGetDir($sVideoFile2)
		EndIf

		If $sContextDir = "" Then
			; Нет контекстной папки - сбрасываем поле
			If $iID = $iInput1 Then
				$sVideoFile1 = ""
				GUICtrlSetData($iID, "")
			ElseIf $iID = $iInput2 Then
				$sVideoFile2 = ""
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


Func _OnEvent_ComboLang()
	Local $sSelected = GUICtrlRead($iComboLang)
	Local $sNewLang = "Russian"
	If $sSelected = "English" Then $sNewLang = "English"
	If $sNewLang = $g_sCurrentLang Then Return
	$g_sCurrentLang = $sNewLang
	$g_sLangFile = @ScriptDir & "\Lang\" & $g_sCurrentLang & ".lng"
	IniWrite($sPathIni, "Settings", "Language", $g_sCurrentLang)
	_LoadLangFile()
	_ApplyLanguage()
EndFunc   ;==>_OnEvent_ComboLang


Func _OnEvent_ComboTheme()
	Local $sSelected = GUICtrlRead($iComboTheme)
	Local $sNew = $g_sTheme
	If $sSelected = Lang("GUI", "ThemeSystem", "System") Then $sNew = "System"
	If $sSelected = Lang("GUI", "ThemeLight", "Light") Then $sNew = "Light"
	If $sSelected = Lang("GUI", "ThemeDark", "Dark") Then $sNew = "Dark"
	If $sNew = $g_sTheme Then Return
	$g_sTheme = $sNew
	IniWrite($sPathIni, "Settings", "Theme", $sNew)
	_ApplyTheme()
EndFunc   ;==>_OnEvent_ComboTheme


Func _UpdateFilesInfo()
	Local $bFile1Exists = FileExists($sVideoFile1)
	Local $bFile2Exists = FileExists($sVideoFile2)

	; Если оба файла существуют, вычисляем crop
	If $bFile1Exists And $bFile2Exists Then
		Local $aInfo1 = _GetVideoInfo($sVideoFile1)
		Local $aInfo2 = _GetVideoInfo($sVideoFile2)
		Local $aCropArgs = _CalculateCropArgs($aInfo1, $aInfo2)

		; Формируем текст для первого файла
		Local $sText1 = Lang("Info", "Resolution", "Resolution") & " " & $aInfo1[0] & "x" & $aInfo1[1]
		If $aCropArgs[2] <> $aInfo1[1] Then
			Local $iCropDiff1 = $aInfo1[1] - $aCropArgs[2]
			$sText1 &= ", " & Lang("Info", "HeightDiff", "height difference") & " " & $iCropDiff1
		EndIf
		GUICtrlSetData($iLabelInfo1, $sText1)

		; Формируем текст для второго файла
		Local $sText2 = Lang("Info", "Resolution", "Resolution") & " " & $aInfo2[0] & "x" & $aInfo2[1]
		If $aCropArgs[3] <> $aInfo2[1] Then
			Local $iCropDiff2 = $aInfo2[1] - $aCropArgs[3]
			$sText2 &= ", " & Lang("Info", "HeightDiff", "height difference") & " " & $iCropDiff2
		EndIf
		GUICtrlSetData($iLabelInfo2, $sText2)

		GUICtrlSetState($iButtonCompare, $GUI_ENABLE)

		; Читаем сдвиг из поля ввода
		Local $sOffset = GUICtrlRead($iInputOffset)
		Local $iOffsetMs = ($sOffset <> "") ? Int($sOffset) : 0

		; Обновляем поле команды
		Local $bIsVertical = (GUICtrlRead($iRadioVertical) = $GUI_CHECKED)
		Local $sCmdLine = _BuildVideoCompareCommand($aInfo1, $aInfo2, $aCropArgs, $bIsVertical, $iOffsetMs)
		GUICtrlSetData($iEditCommand, $sCmdLine)
	Else
		; Показываем только разрешение без crop
		If $bFile1Exists Then
			Local $aInfo1 = _GetVideoInfo($sVideoFile1)
			GUICtrlSetData($iLabelInfo1, Lang("Info", "Resolution", "Resolution") & " " & $aInfo1[0] & "x" & $aInfo1[1])
		Else
			GUICtrlSetData($iLabelInfo1, Lang("GUI", "FileNotSelected", "File not selected"))
		EndIf

		If $bFile2Exists Then
			Local $aInfo2 = _GetVideoInfo($sVideoFile2)
			GUICtrlSetData($iLabelInfo2, Lang("Info", "Resolution", "Resolution") & " " & $aInfo2[0] & "x" & $aInfo2[1])
		Else
			GUICtrlSetData($iLabelInfo2, Lang("GUI", "FileNotSelected", "File not selected"))
		EndIf

		GUICtrlSetState($iButtonCompare, $GUI_DISABLE)
		GUICtrlSetData($iEditCommand, "")
	EndIf
EndFunc   ;==>_UpdateFilesInfo


Func _UpdateCommandField()
	If Not FileExists($sVideoFile1) Or Not FileExists($sVideoFile2) Then
		GUICtrlSetData($iEditCommand, "")
		Return
	EndIf
	Local $aVideo1Info = _GetVideoInfo($sVideoFile1)
	Local $aVideo2Info = _GetVideoInfo($sVideoFile2)
	Local $aCropArgs = _CalculateCropArgs($aVideo1Info, $aVideo2Info)
	Local $sOffset = GUICtrlRead($iInputOffset)
	Local $iOffsetMs = ($sOffset <> "") ? Int($sOffset) : 0
	Local $bIsVertical = (GUICtrlRead($iRadioVertical) = $GUI_CHECKED)
	Local $sCmdLine = _BuildVideoCompareCommand($aVideo1Info, $aVideo2Info, $aCropArgs, $bIsVertical, $iOffsetMs)
	GUICtrlSetData($iEditCommand, $sCmdLine)
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


Func _GetSyncOffset($sFile1, $sFile2)
	Local $iSize1 = FileGetSize($sFile1)
	Local $sMtime1 = FileGetTime($sFile1, $FT_MODIFIED, 1)
	Local $iSize2 = FileGetSize($sFile2)
	Local $sMtime2 = FileGetTime($sFile2, $FT_MODIFIED, 1)
	Local $sMemKey = $sFile1 & "|" & $iSize1 & "|" & $sMtime1 & "|" & $sFile2 & "|" & $iSize2 & "|" & $sMtime2

	; Кеш в памяти
	If $g_oSyncCache.Exists($sMemKey) Then Return $g_oSyncCache.Item($sMemKey)

	; Кеш в памяти (обратная пара — инвертируем сдвиг)
	Local $sMemKeyRev = $sFile2 & "|" & $iSize2 & "|" & $sMtime2 & "|" & $sFile1 & "|" & $iSize1 & "|" & $sMtime1
	If $g_oSyncCache.Exists($sMemKeyRev) Then
		Local $iOffset = -$g_oSyncCache.Item($sMemKeyRev)
		$g_oSyncCache.Item($sMemKey) = $iOffset
		Return $iOffset
	EndIf

	; Кеш на диске (ключ: path1|path2, значение: size1|mtime1|size2|mtime2|offset)
	Local $sDiskKey = $sFile1 & "|" & $sFile2
	Local $sCached = IniRead($sPathCache, "SyncCache", $sDiskKey, "")
	If $sCached <> "" Then
		Local $aParts = StringSplit($sCached, "|")
		If $aParts[0] >= 5 And $aParts[1] = String($iSize1) And $aParts[2] = $sMtime1 _
				And $aParts[3] = String($iSize2) And $aParts[4] = $sMtime2 Then
			Local $iOffset = Int($aParts[5])
			$g_oSyncCache.Item($sMemKey) = $iOffset
			Return $iOffset
		EndIf
	EndIf

	; Кеш на диске (обратная пара — инвертируем сдвиг)
	Local $sDiskKeyRev = $sFile2 & "|" & $sFile1
	Local $sCachedRev = IniRead($sPathCache, "SyncCache", $sDiskKeyRev, "")
	If $sCachedRev <> "" Then
		Local $aParts = StringSplit($sCachedRev, "|")
		If $aParts[0] >= 5 And $aParts[1] = String($iSize2) And $aParts[2] = $sMtime2 _
				And $aParts[3] = String($iSize1) And $aParts[4] = $sMtime1 Then
			Local $iOffset = -Int($aParts[5])
			$g_oSyncCache.Item($sMemKey) = $iOffset
			Return $iOffset
		EndIf
	EndIf

	; Вызов sync.exe с таймаутом и прогрессом
	Local $iOffset = _RunSyncWithTimeout($sFile1, $sFile2, 30)

	; Сохраняем в оба кеша (прямая и обратная пары)
	$g_oSyncCache.Item($sMemKey) = $iOffset
	$g_oSyncCache.Item($sMemKeyRev) = -$iOffset
	IniWrite($sPathCache, "SyncCache", $sDiskKey, $iSize1 & "|" & $sMtime1 & "|" & $iSize2 & "|" & $sMtime2 & "|" & $iOffset)
	IniWrite($sPathCache, "SyncCache", $sDiskKeyRev, $iSize2 & "|" & $sMtime2 & "|" & $iSize1 & "|" & $sMtime1 & "|" & - $iOffset)

	Return $iOffset
EndFunc   ;==>_GetSyncOffset


Func _RunSyncWithTimeout($sFile1, $sFile2, $iTimeoutSec)
	Local $sCmdLine = '"' & $sPathSync & '" sync --v1 "' & $sFile1 & '" --v2 "' & $sFile2 & '"'
	Local $iPid = Run($sCmdLine, "", @SW_HIDE, $STDERR_CHILD + $STDOUT_CHILD)
	Local $hTimer = TimerInit()
	Local $iTimeoutMs = $iTimeoutSec * 1000
	Local $sOutput = ""

	While ProcessExists($iPid)
		Local $iElapsed = TimerDiff($hTimer)

		; Таймаут — убиваем процесс, возвращаем 0
		If $iElapsed >= $iTimeoutMs Then
			ProcessClose($iPid)
			Return 0
		EndIf

		; Обновляем прогрессбар (0–95% за $iTimeoutSec секунд)
		Local $iPct = Int(($iElapsed / $iTimeoutMs) * 95)
		GUICtrlSetData($iProgressSync, $iPct)

		; Читаем stdout неблокирующе
		$sOutput &= StdoutRead($iPid)
		Sleep(100)
	WEnd

	; Дочитываем остатки stdout
	While 1
		Local $sLine = StdoutRead($iPid)
		If @error Then ExitLoop
		$sOutput &= $sLine
	WEnd

	Return Int(StringStripWS($sOutput, 3))
EndFunc   ;==>_RunSyncWithTimeout


Func _TryFillCachedOffset()
	If GUICtrlRead($iCheckSync) <> $GUI_CHECKED Then Return
	If Not FileExists($sVideoFile1) Or Not FileExists($sVideoFile2) Then Return

	Local $iSize1 = FileGetSize($sVideoFile1)
	Local $sMtime1 = FileGetTime($sVideoFile1, $FT_MODIFIED, 1)
	Local $iSize2 = FileGetSize($sVideoFile2)
	Local $sMtime2 = FileGetTime($sVideoFile2, $FT_MODIFIED, 1)
	Local $sMemKey = $sVideoFile1 & "|" & $iSize1 & "|" & $sMtime1 & "|" & $sVideoFile2 & "|" & $iSize2 & "|" & $sMtime2

	; Кеш в памяти
	If $g_oSyncCache.Exists($sMemKey) Then
		GUICtrlSetData($iInputOffset, $g_oSyncCache.Item($sMemKey))
		Return
	EndIf

	; Кеш на диске
	Local $sDiskKey = $sVideoFile1 & "|" & $sVideoFile2
	Local $sCached = IniRead($sPathCache, "SyncCache", $sDiskKey, "")
	If $sCached <> "" Then
		Local $aParts = StringSplit($sCached, "|")
		If $aParts[0] >= 5 And $aParts[1] = String($iSize1) And $aParts[2] = $sMtime1 _
				And $aParts[3] = String($iSize2) And $aParts[4] = $sMtime2 Then
			Local $iOffset = Int($aParts[5])
			$g_oSyncCache.Item($sMemKey) = $iOffset
			GUICtrlSetData($iInputOffset, $iOffset)
		EndIf
	EndIf
EndFunc   ;==>_TryFillCachedOffset


Func _BuildVideoCompareCommand($aVideo1Info, $aVideo2Info, $aCropArgs, $bIsVertical, $iOffsetMs = 0)
	Local $sCmdLine = '"' & $sPathVideoCompare & '"'

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
			'"' & $sVideoFile1 & '" "' & $sVideoFile2 & '"'

	Return $sCmdLine
EndFunc   ;==>_BuildVideoCompareCommand


Func _ShouldUseFullscreen($aVideo1Info, $aVideo2Info, $aCropArgs, $bIsVertical)
	Local $tDesktopRect = _WinAPI_GetWorkArea()
	Local $iDesktopWidth = DllStructGetData($tDesktopRect, "Right")
	Local $iDesktopHeight = DllStructGetData($tDesktopRect, "Bottom")

	Local $iMaxWidth = _Max($aVideo1Info[0], $aVideo2Info[0])
	Local $iMaxHeight = _Max($aCropArgs[2], $aCropArgs[3])

	If $bIsVertical Then
		$iMaxHeight *= 2 ; Вертикальное расположение удваивает высоту
	EndIf

	Return ($iMaxWidth > $iDesktopWidth Or $iMaxHeight > $iDesktopHeight)
EndFunc   ;==>_ShouldUseFullscreen


Func _RunVideoCompare($sCmdLine)
	ConsoleWrite($sCmdLine & @CR)

	Local Const $CREATE_NO_WINDOW = 0x08000000
	Local Const $STARTF_USESTDHANDLES = 0x00000100
	Local Const $HANDLE_FLAG_INHERIT = 0x00000001

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

	DllCall("kernel32.dll", "bool", "CloseHandle", "handle", $hRead)
	DllCall("kernel32.dll", "bool", "CloseHandle", "handle", $hProcess)
	DllCall("kernel32.dll", "bool", "CloseHandle", "handle", DllStructGetData($tPI, "hThread"))

	If StringInStr($sOutput, "Error:") Then MsgBox(48, $sAppName, Lang("Errors", "Error", "Error") & @CR & $sOutput)
EndFunc   ;==>_RunVideoCompare


Func _GetVideoResolution($sVideoPath)
	If $sVideoPath = "" Or Not FileExists($sVideoPath) Then Return ""

	Local $iSize = FileGetSize($sVideoPath)
	Local $sMtime = FileGetTime($sVideoPath, $FT_MODIFIED, 1)
	Local $sMemKey = $sVideoPath & "|" & $iSize & "|" & $sMtime

	; Кеш в памяти
	If $g_oInfoCache.Exists($sMemKey) Then Return $g_oInfoCache.Item($sMemKey)

	; Кеш на диске (формат значения: size|mtime|result)
	Local $sCached = IniRead($sPathCache, "InfoCache", $sVideoPath, "")
	If $sCached <> "" Then
		Local $aParts = StringSplit($sCached, "|")
		If $aParts[0] >= 3 And $aParts[1] = String($iSize) And $aParts[2] = $sMtime Then
			$g_oInfoCache.Item($sMemKey) = $aParts[3]
			Return $aParts[3]
		EndIf
	EndIf

	; Вызов sync.exe info
	Local $sCmdLine = '"' & $sPathSync & '" info "' & $sVideoPath & '"'
	Local $sOutput = _RunToolReadStdout($sCmdLine)
	ConsoleWrite($sVideoPath & " -> " & $sOutput & @CR)

	; Сохраняем в оба кеша
	If $sOutput <> "" Then
		$g_oInfoCache.Item($sMemKey) = $sOutput
		IniWrite($sPathCache, "InfoCache", $sVideoPath, $iSize & "|" & $sMtime & "|" & $sOutput)
	EndIf

	Return $sOutput
EndFunc   ;==>_GetVideoResolution


Func _RunToolReadStdout($sCmdLine)
	Local $iPid = Run($sCmdLine, "", @SW_HIDE, $STDERR_CHILD + $STDOUT_CHILD)
	Local $sOutput = "", $sLine
	While 1
		$sLine = StdoutRead($iPid)
		If @error Then ExitLoop
		$sOutput &= $sLine
	WEnd
	Return StringStripWS($sOutput, 3)
EndFunc   ;==>_RunToolReadStdout


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
	GUICtrlSetResizing($iLogoPic, $iDockStretchH)
	GUICtrlSetResizing($iInput1, $iDockStretchH)
	GUICtrlSetResizing($iLabelInfo1, $iDockStretchH)
	GUICtrlSetResizing($iInput2, $iDockStretchH)
	GUICtrlSetResizing($iLabelInfo2, $iDockStretchH)
	GUICtrlSetResizing($iSeparator, $iDockStretchH)
	GUICtrlSetResizing($iSeparator2, $iDockStretchH_B)
	GUICtrlSetResizing($iEditCommand, $iDockStretchHV)
	GUICtrlSetResizing($iProgressSync, $iDockStretchH)

	; Фиксированные слева
	GUICtrlSetResizing($iLabel1, $iDockFixed)
	GUICtrlSetResizing($iLabelOffset, $iDockFixed)
	GUICtrlSetResizing($iCheckSync, $iDockFixed)
	GUICtrlSetResizing($iLabel2, $iDockFixed)
	GUICtrlSetResizing($iLabelCompare, $iDockFixed)
	GUICtrlSetResizing($iRadioDirect, $iDockFixed)
	GUICtrlSetResizing($iRadioVertical, $iDockFixed)
	GUICtrlSetResizing($iLabelCommand, $iDockFixed)
	GUICtrlSetResizing($iLabelSettings, $iDockFixedBL)
	GUICtrlSetResizing($iLabelLang, $iDockFixedBL)
	GUICtrlSetResizing($iComboLang, $iDockFixedBL)
	GUICtrlSetResizing($iLabelTheme, $iDockFixedBL)
	GUICtrlSetResizing($iComboTheme, $iDockFixedBL)

	; Фиксированные справа
	GUICtrlSetResizing($iButtonChoose1, $iDockFixedRight)
	GUICtrlSetResizing($iButtonSwap, $iDockFixedRight)
	GUICtrlSetResizing($iButtonChoose2, $iDockFixedRight)
	GUICtrlSetResizing($iButtonCompare, $iDockFixedBR)
	GUICtrlSetResizing($iButtonClearCache, $iDockFixedBR)
	GUICtrlSetResizing($iInputOffset, $iDockFixedRight)
EndFunc   ;==>_SetCtrlResizing


Func _RenderHeader()
	_DisposeLogo()

	Local $sLogoPath = @ScriptDir & "\Assets\logo.png"
	If Not FileExists($sLogoPath) Then Return

	_GDIPlus_Startup()
	Local $hImage = _GDIPlus_ImageLoadFromFile($sLogoPath)
	If @error Or Not $hImage Then
		_GDIPlus_Shutdown()
		Return
	EndIf

	Local $iDstW = $iGuiWidth, $iDstH = $iHeaderH
	Local $iSrcW = _GDIPlus_ImageGetWidth($hImage)
	Local $iSrcH = _GDIPlus_ImageGetHeight($hImage)

	; Масштабируем по ширине (cover), обрезаем по высоте
	Local $nScale = $iDstW / $iSrcW
	Local $iScaledW = $iDstW
	Local $iScaledH = Int($iSrcH * $nScale)
	Local $iOffY = Int(($iDstH - $iScaledH) / 2)

	Local $hCanvas = _GDIPlus_BitmapCreateFromScan0($iDstW, $iDstH)
	Local $hGfx = _GDIPlus_ImageGetGraphicsContext($hCanvas)
	_GDIPlus_GraphicsSetInterpolationMode($hGfx, 7) ; HighQualityBicubic
	_GDIPlus_GraphicsSetTextRenderingHint($hGfx, 5) ; AntiAliasGridFit

	; Рисуем логотип (cover: заполняем всю ширину, обрезаем сверху/снизу)
	_GDIPlus_GraphicsDrawImageRect($hGfx, $hImage, 0, $iOffY, $iScaledW, $iScaledH)
	_GDIPlus_ImageDispose($hImage)

	; 4 строки из шпаргалки справа
	Local $aHints[4]
	$aHints[0] = Lang("Hotkeys", "2", "H: Show/hide hints")
	$aHints[1] = Lang("Hotkeys", "3", "V: Show/hide video info")
	$aHints[2] = Lang("Hotkeys", "4", "Space: Play/Pause")
	$aHints[3] = Lang("Hotkeys", "7", "Escape: Exit")

	Local $hFamily = _GDIPlus_FontFamilyCreate($g_sAppFont)
	Local $hFont = _GDIPlus_FontCreate($hFamily, 7.5)
	Local $hBrush = _GDIPlus_BrushCreateSolid(0xCCFFFFFF) ; белый полупрозрачный
	Local $hFormat = _GDIPlus_StringFormatCreate()

	Local $iTextX = Int($iDstW * 0.52)
	Local $iLineH = 14
	Local $iTextY = Int(($iDstH - 4 * $iLineH) / 2)

	For $i = 0 To 3
		Local $tLayout = _GDIPlus_RectFCreate($iTextX, $iTextY + $i * $iLineH, $iDstW - $iTextX - 5, $iLineH)
		_GDIPlus_GraphicsDrawStringEx($hGfx, $aHints[$i], $hFont, $tLayout, $hFormat, $hBrush)
	Next

	_GDIPlus_FontDispose($hFont)
	_GDIPlus_FontFamilyDispose($hFamily)
	_GDIPlus_BrushDispose($hBrush)
	_GDIPlus_StringFormatDispose($hFormat)

	; Устанавливаем в контрол
	Local $hBmp = _GDIPlus_BitmapCreateHBITMAPFromBitmap($hCanvas)
	_GDIPlus_GraphicsDispose($hGfx)
	_GDIPlus_BitmapDispose($hCanvas)
	_GDIPlus_Shutdown()
	If Not $hBmp Then Return

	Local $hWnd = GUICtrlGetHandle($iLogoPic)
	Local Const $STM_SETIMAGE = 0x0172, $IMAGE_BITMAP = 0
	Local $aRet = DllCall("user32.dll", "handle", "SendMessageW", "hwnd", $hWnd, _
			"uint", $STM_SETIMAGE, "wparam", $IMAGE_BITMAP, "lparam", $hBmp)
	If Not @error And IsArray($aRet) And $aRet[0] Then _WinAPI_DeleteObject($aRet[0])
	$hLogoBitmap = $hBmp
EndFunc   ;==>_RenderHeader


Func _SetButtonIcon($iCtrl, $sDllPath, $iIconIndex, $iIconSize)
	; Извлекаем иконку нужного размера
	Local $aIcons = DllCall("user32.dll", "uint", "PrivateExtractIconsW", _
			"wstr", $sDllPath, "int", $iIconIndex, "int", $iIconSize, "int", $iIconSize, _
			"handle*", 0, "uint*", 0, "uint", 1, "uint", 0)
	If @error Or $aIcons[0] = 0 Then Return
	Local $hIcon = $aIcons[5]

	; Создаём ImageList
	Local Const $ILC_COLOR32 = 0x00000020
	Local $aImgList = DllCall("comctl32.dll", "handle", "ImageList_Create", _
			"int", $iIconSize, "int", $iIconSize, "uint", $ILC_COLOR32, "int", 1, "int", 0)
	If @error Then Return
	Local $hImgList = $aImgList[0]

	DllCall("comctl32.dll", "int", "ImageList_ReplaceIcon", _
			"handle", $hImgList, "int", -1, "handle", $hIcon)
	DllCall("user32.dll", "bool", "DestroyIcon", "handle", $hIcon)

	; BUTTON_IMAGELIST: himl, margin(l,t,r,b), uAlign=2 (BUTTON_IMAGELIST_ALIGN_TOP)
	Local $tBIL = DllStructCreate("handle himl;int l;int t;int r;int b;uint uAlign")
	DllStructSetData($tBIL, "himl", $hImgList)
	DllStructSetData($tBIL, "l", 0)
	DllStructSetData($tBIL, "t", 10)
	DllStructSetData($tBIL, "r", 0)
	DllStructSetData($tBIL, "b", 0)
	DllStructSetData($tBIL, "uAlign", 2)

	Local Const $BCM_SETIMAGELIST = 0x1602
	DllCall("user32.dll", "lresult", "SendMessageW", "hwnd", GUICtrlGetHandle($iCtrl), _
			"uint", $BCM_SETIMAGELIST, "wparam", 0, "struct*", $tBIL)
EndFunc   ;==>_SetButtonIcon


Func _DisposeLogo()
	If $hLogoBitmap Then
		_WinAPI_DeleteObject($hLogoBitmap)
		$hLogoBitmap = 0
	EndIf
EndFunc   ;==>_DisposeLogo


Func _RestoreControlsStyle()
	AdlibUnRegister("_RestoreControlsStyle")

	; Восстанавливаем цвет фона input полей
	GUICtrlSetBkColor($iInput1, $g_iClrInput)
	GUICtrlSetBkColor($iInput2, $g_iClrInput)

	; Обновляем информацию о файлах
	_UpdateFilesInfo()
EndFunc   ;==>_RestoreControlsStyle


Func _GetVideoExtensionsFilter()
	Local $sExtensions = ""
	Local $sExt = ''
	For $sExt In $aSupportedExtensions
		$sExtensions &= "*." & $sExt & ";"
	Next
	Return Lang("Filter", "Video", "Video") & " (" & StringTrimRight($sExtensions, 1) & ")"
EndFunc   ;==>_GetVideoExtensionsFilter


Func _SetComboItemHeight($iCtrl, $iItemHeight)
	Local Const $CB_SETITEMHEIGHT = 0x0153
	Local $hWnd = GUICtrlGetHandle($iCtrl)
	If Not $hWnd Then Return
	; wParam = -1 → «поле выбора» комбобокса (selection field)
	DllCall("user32.dll", "lresult", "SendMessageW", "hwnd", $hWnd, "uint", $CB_SETITEMHEIGHT, "wparam", -1, "lparam", $iItemHeight)
EndFunc   ;==>_SetComboItemHeight


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


Func _ApplyAppFont()
	Local $aCandidates[3] = ["Segoe UI", "Tahoma", "MS Shell Dlg 2"]
	For $sCand In $aCandidates
		If _FontExists($sCand) Then
			$g_sAppFont = $sCand
			ExitLoop
		EndIf
	Next
	; MS Shell Dlg 2 — алиас, который система маппит сама, поэтому он всегда доступен как фолбэк
	GUISetFont($g_iAppFontSize, 400, 0, $g_sAppFont, $hGui)
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


Func _IsValidVideoExtension($sFilePath)
	Local $sExt = StringLower(StringRegExpReplace($sFilePath, '^.*\.', ''))
	For $sSupportedExt In $aSupportedExtensions
		If $sExt = $sSupportedExt Then Return True
	Next
	Return False
EndFunc   ;==>_IsValidVideoExtension


Func _SetVideoFile($iInputID, $sFilePath)
	If $iInputID = $iInput1 Then
		$sVideoFile1 = $sFilePath
		GUICtrlSetData($iInput1, $sFilePath = "" ? "" : _GetFileName($sFilePath))
		IniWrite($sPathIni, "LastDirs", "Video1", $sFilePath)
	ElseIf $iInputID = $iInput2 Then
		$sVideoFile2 = $sFilePath
		GUICtrlSetData($iInput2, $sFilePath = "" ? "" : _GetFileName($sFilePath))
		IniWrite($sPathIni, "LastDirs", "Video2", $sFilePath)
	EndIf
	GUICtrlSetData($iInputOffset, "")
	_TryFillCachedOffset()
	_UpdateFilesInfo()
EndFunc   ;==>_SetVideoFile


Func _EnsureIniDefaults()
	If Not FileExists($sPathIni) Then
		; Создаём ini со значениями по умолчанию
		IniWrite($sPathIni, "LastDirs", "Video1", "")
		IniWrite($sPathIni, "LastDirs", "Video2", "")
		IniWrite($sPathIni, "Tools", "VideoCompare", "video-compare.exe")
		IniWrite($sPathIni, "Tools", "Sync", "sync\dist\sync.exe")
	EndIf
EndFunc   ;==>_EnsureIniDefaults


Func _NormalizePath($sValue)
	; Абсолютный путь: диск:\ или UNC \\ или корень /\
	If StringRegExp($sValue, "^(?:[A-Za-z]:\\|\\\\|/)") Then Return $sValue
	; Иначе считаем относительным к папке скрипта
	Return @ScriptDir & "\\" & $sValue
EndFunc   ;==>_NormalizePath


Func _ResolveToolPaths()
	Local $sIniVideoCompare = _NormalizePath(IniRead($sPathIni, "Tools", "VideoCompare", ""))
	If Not FileExists($sPathVideoCompare) And FileExists($sIniVideoCompare) Then
		$sPathVideoCompare = $sIniVideoCompare
	EndIf

	Local $sIniSync = _NormalizePath(IniRead($sPathIni, "Tools", "Sync", ""))
	If Not FileExists($sPathSync) And FileExists($sIniSync) Then
		$sPathSync = $sIniSync
	EndIf
EndFunc   ;==>_ResolveToolPaths


; === Система локализации ===

Func Lang($sSection, $sKey, $sDefault = "")
	If Not IsObj($g_oLangDict) Then Return $sDefault
	Local $sFullKey = $sSection & "." & $sKey
	If $g_oLangDict.Exists($sFullKey) Then Return $g_oLangDict.Item($sFullKey)
	Return $sDefault
EndFunc   ;==>Lang


Func _InitLanguage()
	$g_sCurrentLang = IniRead($sPathIni, "Settings", "Language", "")
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
	$g_oLangDict = ObjCreate("Scripting.Dictionary")
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
			$g_oLangDict($sSection & "." & $sK) = $sV
		EndIf
	Next
EndFunc   ;==>_LoadLangFile


Func _ApplyLanguage()
	GUICtrlSetData($iLabel1, Lang("GUI", "File1", "File 1"))
	GUICtrlSetData($iLabel2, Lang("GUI", "File2", "File 2"))
	GUICtrlSetData($iButtonChoose1, Lang("GUI", "Choose", "Choose"))
	GUICtrlSetData($iButtonChoose2, Lang("GUI", "Choose", "Choose"))
	GUICtrlSetData($iLabelCompare, Lang("GUI", "CompareMode", "Compare:"))
	GUICtrlSetData($iRadioDirect, Lang("GUI", "CompareDirect", "Direct"))
	GUICtrlSetData($iRadioVertical, Lang("GUI", "CompareVertical", "Vertical"))
	GUICtrlSetData($iButtonCompare, Lang("GUI", "Compare", "Compare"))
	GUICtrlSetData($iLabelOffset, Lang("GUI", "Offset", "Offset:"))
	GUICtrlSetData($iCheckSync, Lang("GUI", "SyncVideo", "Synchronize video"))
	GUICtrlSetData($iLabelCommand, Lang("GUI", "TabCommand", "Command"))
	GUICtrlSetData($iLabelSettings, Lang("GUI", "TabSettings", "Settings:"))
	GUICtrlSetData($iLabelLang, Lang("GUI", "Language", "Language:"))
	GUICtrlSetData($iLabelTheme, Lang("GUI", "Theme", "Theme:"))
	_PopulateComboTheme()
	_RenderHeader()
	_UpdateFilesInfo()
	_ApplyTheme()
EndFunc   ;==>_ApplyLanguage


; === Темы оформления ===

Func _InitTheme()
	$g_sTheme = IniRead($sPathIni, "Settings", "Theme", "System")
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
		$g_iClrBg = 0x202020
		$g_iClrFg = 0xE0E0E0
		$g_iClrInfo = 0x9A9A9A
		$g_iClrInput = 0x3C3C3C
		$g_iClrSep = 0x3F3F3F
	Else
		$g_iClrBg = 0xF0F0F0
		$g_iClrFg = 0x000000
		$g_iClrInfo = 0x808080
		$g_iClrInput = 0xFFFFFF
		$g_iClrSep = 0xC0C0C0
	EndIf
EndFunc   ;==>_SetPalette


Func _ApplyTheme()
	Local $bPrevDark = $g_bDarkMode
	_ResolveDarkMode()

	; Применяем UDF-тему ко всему GUI. Не используем _GUIDarkTheme_SwitchTheme —
	; он определяет направление по системной теме, а не по нашему выбору.
	; Важно: _GUIDarkTheme_ApplyLight меняет внутренние UDF-глобалы на светлые значения,
	; но _GUIDarkTheme_ApplyDark их обратно не восстанавливает. Поэтому перед переключением
	; в тёмную тему явно возвращаем тёмную палитру UDF (те же значения, что задаёт _SwitchTheme).
	If Not $g_bThemeInitialized Or $bPrevDark <> $g_bDarkMode Then
		If $g_bThemeInitialized Then __GUIDarkTheme_SubclassCleanup()
		If $g_bDarkMode Then
			$g_iBkColor = 0x1C1C1C
			$COLOR_BG_DARK = 0x121212
			$COLOR_TEXT_LIGHT = 0xE0E0E0
			$COLOR_CONTROL_BG = 0x202020
			$COLOR_BORDER_LIGHT = 0xB0B0B0
			$COLOR_BORDER = 0x3F3F3F
			_GUIDarkTheme_ApplyDark($hGui, True)
		Else
			_GUIDarkTheme_ApplyLight($hGui)
		EndIf
		$g_bThemeInitialized = True
	EndIf

	; Палитра для кастомных элементов
	_SetPalette()

	; UDF ставит фон окна 0x121212 — в тёмной теме делаем его светлее
	If $g_bDarkMode Then GUISetBkColor($g_iClrBg, $hGui)

	; Info-лейблы — серый оттенок
	GUICtrlSetColor($iLabelInfo1, $g_iClrInfo)
	GUICtrlSetColor($iLabelInfo2, $g_iClrInfo)

	; Горизонтальный разделитель
	GUICtrlSetBkColor($iSeparator, $g_iClrSep)
	GUICtrlSetBkColor($iSeparator2, $g_iClrSep)

	; Финальная полная перерисовка окна
	; RDW_INVALIDATE=0x1, RDW_UPDATENOW=0x100, RDW_ALLCHILDREN=0x80
	DllCall("user32.dll", "bool", "RedrawWindow", "hwnd", $hGui, "ptr", 0, "handle", 0, "uint", 0x181)
EndFunc   ;==>_ApplyTheme


Func _PopulateComboTheme()
	Local $sSystem = Lang("GUI", "ThemeSystem", "System")
	Local $sLight = Lang("GUI", "ThemeLight", "Light")
	Local $sDark = Lang("GUI", "ThemeDark", "Dark")
	Local $sCurrent = $sSystem
	If $g_sTheme = "Light" Then $sCurrent = $sLight
	If $g_sTheme = "Dark" Then $sCurrent = $sDark
	GUICtrlSetData($iComboTheme, "")
	GUICtrlSetData($iComboTheme, $sSystem & "|" & $sLight & "|" & $sDark, $sCurrent)
EndFunc   ;==>_PopulateComboTheme
