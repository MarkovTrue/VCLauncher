#pragma compile(Out, VCLauncher.exe)
#pragma compile(Icon, Assets\icon.ico)

#NoTrayIcon
#RequireAdmin

#include <GUIConstantsEx.au3>
#include <WindowsConstants.au3>
#include <FileConstants.au3>
#include <EditConstants.au3>
#include <StaticConstants.au3>
#include <Array.au3>
#include <WinAPI.au3>
#include <Math.au3>

Opt("GUIOnEventMode", 1)

; Имя ini-файла для хранения последних путей
Global $sAppName = "VCLauncher 0.25"
Global $sPathIni = @ScriptDir & '\VCLauncher.ini'
Global $sPathFfprobe = @ScriptDir & '\ffprobe.exe'
Global $sPathVideoCompare = @ScriptDir & '\video-compare.exe'

Global $iGuiWidth = 460
Global $iGuiHeight = 120

; Создание ini по умолчанию, если отсутствует
_EnsureIniDefaults()

; Ссылки на элементы GUI
Global $hGui, $iInput1, $iInput2, $iButtonChoose1, $iButtonChoose2, $iButtonCompare, $iRadioDirect, $iRadioVertical
; Остальные переменные (чтение и нормализация путей; относительные пути считаем от папки скрипта)
Global $sVideoFile1 = _NormalizePath(IniRead($sPathIni, "LastDirs", "Video1", ""))
Global $sVideoFile2 = _NormalizePath(IniRead($sPathIni, "LastDirs", "Video2", ""))


_ResolveToolPaths()

If Not FileExists($sPathVideoCompare) Then
	MsgBox(48, $sAppName, 'Не найден "video-compare.exe". Укажите путь в секции [Tools] -> VideoCompare в ' & $sPathIni & @CR & 'Текущий путь: ' & $sPathVideoCompare)
	Exit
EndIf

If Not FileExists($sPathFfprobe) Then
	MsgBox(48, $sAppName, 'Не найден "ffprobe.exe". Укажите путь в секции [Tools] -> FFprobe в ' & $sPathIni & @CR & 'Текущий путь: ' & $sPathFfprobe)
	Exit
EndIf

_MainGUI()
_DefineEvents()

While 1
	Sleep(50)
WEnd


Func _MainGUI()
	$hGui = GUICreate($sAppName, $iGuiWidth, $iGuiHeight, -1, -1, $WS_SIZEBOX + $WS_SYSMENU + $WS_MINIMIZEBOX)

	Local $iLabel1 = GUICtrlCreateLabel("Файл 1", 10, 10, 55, 20)
	GUICtrlSetResizing($iLabel1, $GUI_DOCKLEFT + $GUI_DOCKTOP + $GUI_DOCKWIDTH + $GUI_DOCKHEIGHT)

	$iInput1 = GUICtrlCreateInput("", 60, 8, $iGuiWidth - 150, 20)
;~ GUICtrlSetState(-1, $GUI_DROPACCEPTED)
	GUICtrlSetResizing($iInput1, $GUI_DOCKLEFT + $GUI_DOCKTOP + $GUI_DOCKRIGHT + $GUI_DOCKHEIGHT)
	If $sVideoFile1 <> "" Then GUICtrlSetData($iInput1, _GetFileName($sVideoFile1))

	$iButtonChoose1 = GUICtrlCreateButton("Выбрать", $iGuiWidth - 82, 7, 74, 22)
	GUICtrlSetResizing($iButtonChoose1, $GUI_DOCKRIGHT + $GUI_DOCKTOP + $GUI_DOCKWIDTH + $GUI_DOCKHEIGHT)

	Local $iLabel2 = GUICtrlCreateLabel("Файл 2", 10, 38, 55, 20)
	GUICtrlSetResizing($iLabel2, $GUI_DOCKLEFT + $GUI_DOCKTOP + $GUI_DOCKWIDTH + $GUI_DOCKHEIGHT)

	$iInput2 = GUICtrlCreateInput("", 60, 36, $iGuiWidth - 150, 20)
;~ GUICtrlSetState(-1, $GUI_DROPACCEPTED)
	GUICtrlSetResizing($iInput2, $GUI_DOCKLEFT + $GUI_DOCKTOP + $GUI_DOCKRIGHT + $GUI_DOCKHEIGHT)
	If $sVideoFile2 <> "" Then GUICtrlSetData($iInput2, _GetFileName($sVideoFile2))

	$iButtonChoose2 = GUICtrlCreateButton("Выбрать", $iGuiWidth - 82, 35, 74, 22)
	GUICtrlSetResizing($iButtonChoose2, $GUI_DOCKRIGHT + $GUI_DOCKTOP + $GUI_DOCKWIDTH + $GUI_DOCKHEIGHT)

	; Радио-кнопки для режима сравнения
	$iRadioDirect = GUICtrlCreateRadio("Прямое сравнение", 60, 68, 130, 20)
	GUICtrlSetResizing($iRadioDirect, $GUI_DOCKLEFT + $GUI_DOCKTOP + $GUI_DOCKWIDTH + $GUI_DOCKHEIGHT)
	$iRadioVertical = GUICtrlCreateRadio("Вертикальное сравнение", 190, 68, 160, 20)
	GUICtrlSetResizing($iRadioVertical, $GUI_DOCKLEFT + $GUI_DOCKTOP + $GUI_DOCKWIDTH + $GUI_DOCKHEIGHT)
	GUICtrlSetState($iRadioDirect, $GUI_CHECKED) ; По умолчанию включено "Прямое сравнение"

	$iButtonCompare = GUICtrlCreateButton("Сравнить", $iGuiWidth - 82, 66, 74, 25)
	GUICtrlSetResizing($iButtonCompare, $GUI_DOCKRIGHT + $GUI_DOCKTOP + $GUI_DOCKWIDTH + $GUI_DOCKHEIGHT)

	; Проверка наличия файлов при запуске
	_ValidateFiles()

	GUISetState(@SW_SHOW)
EndFunc   ;==>_MainGUI


Func _DefineEvents()
	GUICtrlSetOnEvent($iButtonChoose1, "_OnEvent_ButtonChoose")
	GUICtrlSetOnEvent($iButtonChoose2, "_OnEvent_ButtonChoose")
	GUICtrlSetOnEvent($iButtonCompare, "_OnEvent_ButtonCompare")
	GUISetOnEvent($GUI_EVENT_CLOSE, "_OnEvent_GUI_EVENT_CLOSE")

	; Ограничение на высоту окна через WM_GETMINMAXINFO
	GUIRegisterMsg($WM_GETMINMAXINFO, "_OnEvent_WM_GETMINMAXINFO")
	; Подписываемся, чтобы фильтровать символы в поле ввода
	GUIRegisterMsg($WM_COMMAND, "_OnEvent_WM_COMMAND")
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
		DllStructSetData($tMINMAXINFO, "MaxTrackSizeY", $iGuiHeight + 14) ; максимальная высота окна
	EndIf
	Return $GUI_RUNDEFMSG
EndFunc   ;==>_OnEvent_WM_GETMINMAXINFO


Func _OnEvent_ButtonChoose()
	Local $iButtonID = @GUI_CtrlId
	Local $sTitle, $sIniKey, $sInputCtrl

	If $iButtonID = $iButtonChoose1 Then
		$sTitle = "Выберите видео 1"
		$sIniKey = "Video1"
		$sInputCtrl = $iInput1
		Local $sCurrentFile = $sVideoFile1
	Else
		$sTitle = "Выберите видео 2"
		$sIniKey = "Video2"
		$sInputCtrl = $iInput2
		Local $sCurrentFile = $sVideoFile2
	EndIf

	Local $sFile = FileOpenDialog($sTitle, _PathGetDir($sCurrentFile), "Видео (*.mp4;*.avi;*.mkv;*.mov;*.wmv)|Все файлы (*.*)", 1, _GetFileName($sCurrentFile))
	If Not @error And FileExists($sFile) Then
		If $iButtonID = $iButtonChoose1 Then
			$sVideoFile1 = $sFile
		Else
			$sVideoFile2 = $sFile
		EndIf
		GUICtrlSetData($sInputCtrl, _GetFileName($sFile))
		IniWrite($sPathIni, "LastDirs", $sIniKey, $sFile)
		_ValidateFiles()
	EndIf
EndFunc   ;==>_OnEvent_ButtonChoose


Func _ValidateFiles()
	If FileExists($sVideoFile1) And FileExists($sVideoFile2) Then
		GUICtrlSetState($iButtonCompare, $GUI_ENABLE)
	Else
		GUICtrlSetState($iButtonCompare, $GUI_DISABLE)
	EndIf
EndFunc   ;==>_ValidateFiles


Func _OnEvent_ButtonCompare()
	; Строки и управляющие ключи
	Local $sCropArgsLeft = "", $sCropArgsRight = "", $sCmdLine = ""

	; Размер рабочей области экрана
	Local $tDesktopRect = _WinAPI_GetWorkArea()
	Local $iDesktopWidth = DllStructGetData($tDesktopRect, "Right")
	Local $iDesktopHeight = DllStructGetData($tDesktopRect, "Bottom")

	; Получение разрешений видео
	Local $sVideoRes1 = _GetVideoResolution($sVideoFile1)
	Local $sVideoRes2 = _GetVideoResolution($sVideoFile2)

	Local $iVideo1Width = _GetWidthFromFFprobeString($sVideoRes1)
	Local $iVideo1Height = _GetHeightFromFFprobeString($sVideoRes1)
	Local $iVideo2Width = _GetWidthFromFFprobeString($sVideoRes2)
	Local $iVideo2Height = _GetHeightFromFFprobeString($sVideoRes2)

	; Высоты после возможного crop
	Local $iCropVideo1Height = $iVideo1Height
	Local $iCropVideo2Height = $iVideo2Height

	; Выравнивание по ширине и crop по высоте
	If $iVideo1Width <> $iVideo2Width Or $iVideo1Height <> $iVideo2Height Then

		; Общая целевая ширина
		Local $iTargetWidth = _Min($iVideo1Width, $iVideo2Width)

		; Масштаб
		Local $nScale1 = $iTargetWidth / $iVideo1Width
		Local $nScale2 = $iTargetWidth / $iVideo2Width

		; Высоты после масштабирования
		Local $iScaledH1 = Round($iVideo1Height * $nScale1)
		Local $iScaledH2 = Round($iVideo2Height * $nScale2)

		; Разница высот
		Local $iDiffHeight = Abs($iScaledH1 - $iScaledH2)

		If $iDiffHeight > 0 Then
			If $iScaledH1 > $iScaledH2 Then
				; Crop видео 1
				Local $iOrigDiff = Round($iDiffHeight / $nScale1)
				$sCropArgsLeft = "-l crop=iw:ih-" & $iOrigDiff & " "
				$iCropVideo1Height -= $iOrigDiff
			Else
				; Crop видео 2
				Local $iOrigDiff = Round($iDiffHeight / $nScale2)
				$sCropArgsRight = "-r crop=iw:ih-" & $iOrigDiff & " "
				$iCropVideo2Height -= $iOrigDiff
			EndIf
		EndIf
	EndIf

	; Формирование команды video-compare
	$sCmdLine = '"' & $sPathVideoCompare & '"'

	; Режим сравнения
	Local $iMaxVideoWidth, $iMaxVideoHeight

	If GUICtrlRead($iRadioVertical) = $GUI_CHECKED Then
		; Вертикальное сравнение
		$iMaxVideoWidth = _Max($iVideo1Width, $iVideo2Width)
		$iMaxVideoHeight = _Max($iCropVideo1Height, $iCropVideo2Height) * 2

		If $iMaxVideoWidth > $iDesktopWidth Or $iMaxVideoHeight > $iDesktopHeight Then
			$sCmdLine &= " -W"
		EndIf

		$sCmdLine &= " -m vstack"
	Else
		; Горизонтальное сравнение
		$iMaxVideoWidth = _Max($iVideo1Width, $iVideo2Width)
		$iMaxVideoHeight = _Max($iCropVideo1Height, $iCropVideo2Height)

		If $iMaxVideoWidth > $iDesktopWidth Or $iMaxVideoHeight > $iDesktopHeight Then
			$sCmdLine &= " -W"
		EndIf
	EndIf

	; Добавляем crop и пути к файлам
	$sCmdLine &= " " & $sCropArgsLeft & $sCropArgsRight & _
			'"' & $sVideoFile1 & '" "' & $sVideoFile2 & '"'

	; Запуск сравнения
	_RunVideoCompare($sCmdLine)

EndFunc   ;==>_OnEvent_ButtonCompare


Func _RunVideoCompare($sCmdLine)
	ConsoleWrite($sCmdLine & @CR)

	Local $sOutput = "", $sLine
	; Выполняем
	Local $iPid = Run($sCmdLine)
	; Читаем вывод, как чек возможных ошибок
	While 1
		$sLine = StdoutRead($iPid)
		If @error Then ExitLoop
		$sOutput &= $sLine
	WEnd

	If StringInStr($sOutput, "Error:") Then MsgBox(48, $sAppName, 'Ошибка' & @CR & $sOutput)
EndFunc   ;==>_RunVideoCompare


Func _GetVideoResolution($sVideoPath)
	Local $sOutput = "", $sLine
	; Команда ffprobe
	Local $sCmdLine = '"' & $sPathFfprobe & '" -v error -select_streams v:0 -show_entries stream=width,height -of csv=p=0 "' & $sVideoPath & '"'
	; Выполняем
	Local $iPid = Run($sCmdLine, "", @SW_HIDE, $STDERR_CHILD + $STDOUT_CHILD)
	; Читаем вывод
	While 1
		$sLine = StdoutRead($iPid)
		If @error Then ExitLoop
		$sOutput &= $sLine
	WEnd

	ConsoleWrite(@ComSpec & " /c " & $sCmdLine & @CR)
;~ 	MsgBox(48, $sAppName, $sVideoPath & @CR & $sOutput)
	ConsoleWrite($sVideoPath & @CR & $sOutput & @CR)

	Return $sOutput
EndFunc   ;==>_GetVideoResolution


Func _GetWidthFromFFprobeString($sVar)
	Local $aLineSplit = StringSplit($sVar, ',')
	If $aLineSplit[0] >= 2 Then
		Return Int($aLineSplit[1])
	EndIf
;~ 	MsgBox(48, $sAppName, '_GetWidthFromFFprobeString ошибка' & @CR & $sVar)
	Return 0
EndFunc   ;==>_GetWidthFromFFprobeString


Func _GetHeightFromFFprobeString($sVar)
	Local $aLineSplit = StringSplit($sVar, ',')
	If $aLineSplit[0] >= 2 Then
		Return Int($aLineSplit[2])
	EndIf
;~ 	MsgBox(48, $sAppName, '_GetHeightFromFFprobeString ошибка' & @CR & $sVar)
	Return 0
EndFunc   ;==>_GetHeightFromFFprobeString


Func _OnEvent_GUI_EVENT_CLOSE()
	Exit
EndFunc   ;==>_OnEvent_GUI_EVENT_CLOSE


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


Func _EnsureIniDefaults()
	If Not FileExists($sPathIni) Then
		; Последовательно создаём ini со значениями по умолчанию, без переменных окружения
		IniWrite($sPathIni, "LastDirs", "Video1", "")
		IniWrite($sPathIni, "LastDirs", "Video2", "")
		IniWrite($sPathIni, "Tools", "FFprobe", "ffprobe.exe")
		IniWrite($sPathIni, "Tools", "VideoCompare", "video-compare.exe")
	EndIf
EndFunc   ;==>_EnsureIniDefaults


Func _NormalizePath($sValue)
	; Абсолютный путь: диск:\ или UNC \\ или корень /\
	If StringRegExp($sValue, "^(?:[A-Za-z]:\\|\\\\|/)") Then Return $sValue
	; Иначе считаем относительным к папке скрипта
	Return @ScriptDir & "\\" & $sValue
EndFunc   ;==>_NormalizePath


Func _ResolveToolPaths()
	Local $sIniFfprobe = _NormalizePath(IniRead($sPathIni, "Tools", "FFprobe", ""))
	If Not FileExists($sPathFfprobe) And FileExists($sIniFfprobe) Then
		$sPathFfprobe = $sIniFfprobe
	EndIf

	Local $sIniVideoCompare = _NormalizePath(IniRead($sPathIni, "Tools", "VideoCompare", ""))
	If Not FileExists($sPathVideoCompare) And FileExists($sIniVideoCompare) Then
		$sPathVideoCompare = $sIniVideoCompare
	EndIf
EndFunc   ;==>_ResolveToolPaths


Func _OnEvent_WM_COMMAND($hWnd, $iMsg, $wParam, $lParam)
	#forceref $hWnd, $iMsg, $lParam
;~	Local $nNotify = BitShift($wParam, 16) ; Код уведомления
	Local $nID = BitAND($wParam, 0xFFFF) ; ID элемента управления

	If $nID = $iInput1 Or $nID = $iInput2 Then
		; Проверяем, что сообщение пришло от нужного поля
;~ 		If $nNotify = $EN_UPDATE Then
		Local $sInputValue = GUICtrlRead($nID)
		If $sInputValue = "" Then Return

		Local $sFullPath = _NormalizePath($sInputValue)

		If FileExists($sFullPath) Then
			If $nID = $iInput1 Then
				$sVideoFile1 = $sFullPath
				GUICtrlSetData($nID, _GetFileName($sFullPath))
				IniWrite($sPathIni, "LastDirs", "Video1", $sFullPath)
			EndIf

			If $nID = $iInput2 Then
				$sVideoFile2 = $sFullPath
				GUICtrlSetData($nID, _GetFileName($sFullPath))
				IniWrite($sPathIni, "LastDirs", "Video2", $sFullPath)
			EndIf
		EndIf

		_ValidateFiles()
;~ 		EndIf
	EndIf

	Return $GUI_RUNDEFMSG
EndFunc   ;==>_OnEvent_WM_COMMAND
