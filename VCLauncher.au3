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

Opt("GUIOnEventMode", 1)

; Константы приложения
Global Const $sAppName = "VCLauncher 1.01"
Global Const $sPathIni = @ScriptDir & '\VCLauncher.ini'
Global Const $aSupportedExtensions[] = ["mp4", "avi", "mkv", "mov", "wmv", "webm", "mpg", "mpeg"]
Global Const $iGuiWidth = 460
Global Const $iGuiHeight = 152
Global Const $EN_KILLFOCUS = 0x0200

; Переменные путей к инструментам (загружаются из ini)
Global $sPathFfprobe = @ScriptDir & '\ffprobe.exe'
Global $sPathVideoCompare = @ScriptDir & '\video-compare.exe'

; Создание ini по умолчанию, если отсутствует
_EnsureIniDefaults()

; Ссылки на элементы GUI
Global $hGui, $iInput1, $iInput2, $iButtonChoose1, $iButtonChoose2, $iButtonCompare, $iRadioDirect, $iRadioVertical
Global $iLabelInfo1, $iLabelInfo2
; Остальные переменные (чтение и нормализация путей; относительные пути считаем от папки скрипта)
Global $sVideoFile1 = _NormalizePath(IniRead($sPathIni, "LastDirs", "Video1", ""))
Global $sVideoFile2 = _NormalizePath(IniRead($sPathIni, "LastDirs", "Video2", ""))


_ResolveToolPaths()

If Not FileExists($sPathVideoCompare) Then
	MsgBox(48, $sAppName, 'Не найден "video-compare.exe". Укажите полный путь в файле настроек, в секции Tools.' & @CR & @CR & _
			'Сейчас будет открыт файл настроек.')
	ShellExecute($sPathIni)
	Exit
EndIf

If Not FileExists($sPathFfprobe) Then
	MsgBox(48, $sAppName, 'Не найден "ffprobe.exe". Укажите полный путь в файле настроек, в секции Tools.' & @CR & @CR & _
			'Сейчас будет открыт файл настроек.')
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

	Local $iLabel1 = GUICtrlCreateLabel("Файл 1", 10, 10, 55, 20)
	GUICtrlSetResizing($iLabel1, $GUI_DOCKLEFT + $GUI_DOCKTOP + $GUI_DOCKWIDTH + $GUI_DOCKHEIGHT)

	$iInput1 = GUICtrlCreateInput("", 60, 8, $iGuiWidth - 150, 20)
	GUICtrlSetResizing($iInput1, $GUI_DOCKLEFT + $GUI_DOCKTOP + $GUI_DOCKRIGHT + $GUI_DOCKHEIGHT)
	GUICtrlSetState($iInput1, $GUI_DROPACCEPTED)
	If $sVideoFile1 <> "" Then GUICtrlSetData($iInput1, _GetFileName($sVideoFile1))

	$iButtonChoose1 = GUICtrlCreateButton("Выбрать", $iGuiWidth - 82, 7, 74, 22)
	GUICtrlSetResizing($iButtonChoose1, $GUI_DOCKRIGHT + $GUI_DOCKTOP + $GUI_DOCKWIDTH + $GUI_DOCKHEIGHT)

	$iLabelInfo1 = GUICtrlCreateLabel("Файл не выбран", 61, 30, $iGuiWidth - 150, 15)
	GUICtrlSetResizing($iLabelInfo1, $GUI_DOCKLEFT + $GUI_DOCKTOP + $GUI_DOCKRIGHT + $GUI_DOCKHEIGHT)
	GUICtrlSetColor($iLabelInfo1, 0x808080)

	Local $iLabel2 = GUICtrlCreateLabel("Файл 2", 10, 52, 55, 20)
	GUICtrlSetResizing($iLabel2, $GUI_DOCKLEFT + $GUI_DOCKTOP + $GUI_DOCKWIDTH + $GUI_DOCKHEIGHT)

	$iInput2 = GUICtrlCreateInput("", 60, 50, $iGuiWidth - 150, 20)
	GUICtrlSetResizing($iInput2, $GUI_DOCKLEFT + $GUI_DOCKTOP + $GUI_DOCKRIGHT + $GUI_DOCKHEIGHT)
	GUICtrlSetState($iInput2, $GUI_DROPACCEPTED)
	If $sVideoFile2 <> "" Then GUICtrlSetData($iInput2, _GetFileName($sVideoFile2))

	$iButtonChoose2 = GUICtrlCreateButton("Выбрать", $iGuiWidth - 82, 49, 74, 22)
	GUICtrlSetResizing($iButtonChoose2, $GUI_DOCKRIGHT + $GUI_DOCKTOP + $GUI_DOCKWIDTH + $GUI_DOCKHEIGHT)

	$iLabelInfo2 = GUICtrlCreateLabel("Файл не выбран", 61, 72, $iGuiWidth - 150, 15)
	GUICtrlSetResizing($iLabelInfo2, $GUI_DOCKLEFT + $GUI_DOCKTOP + $GUI_DOCKRIGHT + $GUI_DOCKHEIGHT)
	GUICtrlSetColor($iLabelInfo2, 0x808080)

	; Горизонтальная разделительная линия
	Local $iSeparator = GUICtrlCreateLabel("", 0, 91, $iGuiWidth, 1)
	GUICtrlSetResizing($iSeparator, $GUI_DOCKLEFT + $GUI_DOCKTOP + $GUI_DOCKRIGHT + $GUI_DOCKHEIGHT)
	GUICtrlSetBkColor($iSeparator, 0xC0C0C0)

	; Радио-кнопки для режима сравнения
	$iRadioDirect = GUICtrlCreateRadio("Прямое сравнение", 60, 101, 130, 20)
	GUICtrlSetResizing($iRadioDirect, $GUI_DOCKLEFT + $GUI_DOCKTOP + $GUI_DOCKWIDTH + $GUI_DOCKHEIGHT)
	$iRadioVertical = GUICtrlCreateRadio("Вертикальное сравнение", 190, 101, 160, 20)
	GUICtrlSetResizing($iRadioVertical, $GUI_DOCKLEFT + $GUI_DOCKTOP + $GUI_DOCKWIDTH + $GUI_DOCKHEIGHT)
	GUICtrlSetState($iRadioDirect, $GUI_CHECKED) ; По умолчанию включено "Прямое сравнение"

	$iButtonCompare = GUICtrlCreateButton("Сравнить", $iGuiWidth - 82, 99, 74, 25)
	GUICtrlSetResizing($iButtonCompare, $GUI_DOCKRIGHT + $GUI_DOCKTOP + $GUI_DOCKWIDTH + $GUI_DOCKHEIGHT)

	; Проверка файлов и обновление информации при запуске
	_UpdateFilesInfo()

	GUISetState(@SW_SHOW)
EndFunc   ;==>_MainGUI


Func _DefineEvents()
	GUICtrlSetOnEvent($iButtonChoose1, "_OnEvent_ButtonChoose")
	GUICtrlSetOnEvent($iButtonChoose2, "_OnEvent_ButtonChoose")
	GUICtrlSetOnEvent($iButtonCompare, "_OnEvent_ButtonCompare")
	GUISetOnEvent($GUI_EVENT_CLOSE, "_OnEvent_GUI_EVENT_CLOSE")
	GUISetOnEvent($GUI_EVENT_DROPPED, "_OnEvent_GUI_EVENT_DROPPED")

	; Ограничение высоты окна через WM_GETMINMAXINFO
	GUIRegisterMsg($WM_GETMINMAXINFO, "_OnEvent_WM_GETMINMAXINFO")
	; Обработка ввода путей в текстовые поля
	GUIRegisterMsg($WM_COMMAND, "_OnEvent_WM_COMMAND")
	; Обработка drag-and-drop для визуальной реакции
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

	Local $sFile = FileOpenDialog($sTitle, _PathGetDir($sCurrentFile), _GetVideoExtensionsFilter(), 1, _GetFileName($sCurrentFile))
	If Not @error And FileExists($sFile) Then
		If $iButtonID = $iButtonChoose1 Then
			$sVideoFile1 = $sFile
		Else
			$sVideoFile2 = $sFile
		EndIf
		GUICtrlSetData($sInputCtrl, _GetFileName($sFile))
		IniWrite($sPathIni, "LastDirs", $sIniKey, $sFile)
		_UpdateFilesInfo()
	EndIf
EndFunc   ;==>_OnEvent_ButtonChoose


; Обновляет информацию о файлах (разрешение + crop) и валидирует доступность кнопки сравнения
Func _UpdateFilesInfo()
	Local $bFile1Exists = FileExists($sVideoFile1)
	Local $bFile2Exists = FileExists($sVideoFile2)

	; Если оба файла существуют, вычисляем crop
	If $bFile1Exists And $bFile2Exists Then
		Local $aInfo1 = _GetVideoInfo($sVideoFile1)
		Local $aInfo2 = _GetVideoInfo($sVideoFile2)
		Local $aCropArgs = _CalculateCropArgs($aInfo1, $aInfo2)

		; Формируем текст для первого файла
		Local $sText1 = "Разрешение " & $aInfo1[0] & "x" & $aInfo1[1]
		If $aCropArgs[2] <> $aInfo1[1] Then
			Local $iCropDiff1 = $aInfo1[1] - $aCropArgs[2]
			$sText1 &= ", разница высоты " & $iCropDiff1
		EndIf
		GUICtrlSetData($iLabelInfo1, $sText1)

		; Формируем текст для второго файла
		Local $sText2 = "Разрешение " & $aInfo2[0] & "x" & $aInfo2[1]
		If $aCropArgs[3] <> $aInfo2[1] Then
			Local $iCropDiff2 = $aInfo2[1] - $aCropArgs[3]
			$sText2 &= ", разница высоты " & $iCropDiff2
		EndIf
		GUICtrlSetData($iLabelInfo2, $sText2)

		GUICtrlSetState($iButtonCompare, $GUI_ENABLE)
	Else
		; Показываем только разрешение без crop
		If $bFile1Exists Then
			Local $aInfo1 = _GetVideoInfo($sVideoFile1)
			GUICtrlSetData($iLabelInfo1, "Разрешение " & $aInfo1[0] & "x" & $aInfo1[1])
		Else
			GUICtrlSetData($iLabelInfo1, "Файл не выбран")
		EndIf

		If $bFile2Exists Then
			Local $aInfo2 = _GetVideoInfo($sVideoFile2)
			GUICtrlSetData($iLabelInfo2, "Разрешение " & $aInfo2[0] & "x" & $aInfo2[1])
		Else
			GUICtrlSetData($iLabelInfo2, "Файл не выбран")
		EndIf

		GUICtrlSetState($iButtonCompare, $GUI_DISABLE)
	EndIf
EndFunc   ;==>_UpdateFilesInfo


Func _OnEvent_ButtonCompare()
	; Получаем разрешения видео
	Local $aVideo1Info = _GetVideoInfo($sVideoFile1)
	Local $aVideo2Info = _GetVideoInfo($sVideoFile2)

	; Вычисляем параметры crop для выравнивания
	Local $aCropArgs = _CalculateCropArgs($aVideo1Info, $aVideo2Info)

	; Формируем и выполняем команду
	Local $bIsVertical = (GUICtrlRead($iRadioVertical) = $GUI_CHECKED)
	Local $sCmdLine = _BuildVideoCompareCommand($aVideo1Info, $aVideo2Info, $aCropArgs, $bIsVertical)
	_RunVideoCompare($sCmdLine)
EndFunc   ;==>_OnEvent_ButtonCompare


; Получает разрешение видео и возвращает массив [width, height]
Func _GetVideoInfo($sVideoPath)
	Local $sResolution = _GetVideoResolution($sVideoPath)
	Local $aInfo[2]
	$aInfo[0] = _GetWidthFromFFprobeString($sResolution)
	$aInfo[1] = _GetHeightFromFFprobeString($sResolution)
	Return $aInfo
EndFunc   ;==>_GetVideoInfo


; Рассчитывает параметры crop для выравнивания видео по пропорциям
; Возвращает массив [cropLeft, cropRight, height1, height2]
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


; Формирует командную строку для video-compare
Func _BuildVideoCompareCommand($aVideo1Info, $aVideo2Info, $aCropArgs, $bIsVertical)
	Local $sCmdLine = '"' & $sPathVideoCompare & '"'

	; Проверяем, нужно ли окно на весь экран
	If _ShouldUseFullscreen($aVideo1Info, $aVideo2Info, $aCropArgs, $bIsVertical) Then
		$sCmdLine &= " -W"
	EndIf

	; Режим сравнения
	If $bIsVertical Then
		$sCmdLine &= " -m vstack"
	EndIf

	; Добавляем crop и пути к файлам
	$sCmdLine &= " " & $aCropArgs[0] & $aCropArgs[1] & _
			'"' & $sVideoFile1 & '" "' & $sVideoFile2 & '"'

	Return $sCmdLine
EndFunc   ;==>_BuildVideoCompareCommand


; Проверяет, нужен ли полноэкранный режим
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

	Local $sOutput = "", $sLine
	; Выполняем
	Local $iPid = Run($sCmdLine)
	; Читаем вывод для проверки ошибок
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


; Обработка WM_DROPFILES для визуальной реакции при drag-and-drop
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


; Восстанавливает исходный стиль элементов после drag-and-drop
Func _RestoreControlsStyle()
	AdlibUnRegister("_RestoreControlsStyle")

	; Восстанавливаем цвет фона input полей
	GUICtrlSetBkColor($iInput1, $COLOR_WHITE)
	GUICtrlSetBkColor($iInput2, $COLOR_WHITE)

	; Обновляем информацию о файлах
	_UpdateFilesInfo()
EndFunc   ;==>_RestoreControlsStyle


; Обрабатывает drag-and-drop файлов
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
		MsgBox(48, $sAppName, 'Неподдерживаемый формат файла: ' & $sExt & @CR & @CR & 'Поддерживаемые форматы: ' & _GetVideoExtensionsFilter())
		Return
	EndIf

	; Обновляем данные
	_SetVideoFile($iDropId, $sDropFile)
EndFunc   ;==>_OnEvent_GUI_EVENT_DROPPED


; Формирует фильтр расширений для FileOpenDialog
Func _GetVideoExtensionsFilter()
	Local $sExtensions = ""
	Local $sExt = ''
	For $sExt In $aSupportedExtensions
		$sExtensions &= "*." & $sExt & ";"
	Next
	Return "Видео (" & StringTrimRight($sExtensions, 1) & ")"
EndFunc   ;==>_GetVideoExtensionsFilter


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


; Проверяет, является ли расширение файла поддерживаемым
Func _IsValidVideoExtension($sFilePath)
	Local $sExt = StringLower(StringRegExpReplace($sFilePath, '^.*\.', ''))
	For $sSupportedExt In $aSupportedExtensions
		If $sExt = $sSupportedExt Then Return True
	Next
	Return False
EndFunc   ;==>_IsValidVideoExtension


; Обновляет видеофайл для указанного поля ввода
Func _SetVideoFile($nInputID, $sFilePath)
	If $nInputID = $iInput1 Then
		$sVideoFile1 = $sFilePath
		GUICtrlSetData($iInput1, $sFilePath = "" ? "" : _GetFileName($sFilePath))
		IniWrite($sPathIni, "LastDirs", "Video1", $sFilePath)
	ElseIf $nInputID = $iInput2 Then
		$sVideoFile2 = $sFilePath
		GUICtrlSetData($iInput2, $sFilePath = "" ? "" : _GetFileName($sFilePath))
		IniWrite($sPathIni, "LastDirs", "Video2", $sFilePath)
	EndIf
	_UpdateFilesInfo()
EndFunc   ;==>_SetVideoFile


Func _EnsureIniDefaults()
	If Not FileExists($sPathIni) Then
		; Создаём ini со значениями по умолчанию
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
	Local $nID = BitAND($wParam, 0xFFFF)
	Local $nNotifyCode = BitShift($wParam, 16)

	; Обрабатываем только событие EN_KILLFOCUS - когда пользователь завершил редактирование
	If $nNotifyCode <> $EN_KILLFOCUS Then Return $GUI_RUNDEFMSG

	If $nID <> $iInput1 And $nID <> $iInput2 Then Return $GUI_RUNDEFMSG

	Local $sInputValue = GUICtrlRead($nID)
	If $sInputValue = "" Then Return $GUI_RUNDEFMSG

	Local $sFullPath = ""

	; Проверяем, является ли путь абсолютным
	If StringRegExp($sInputValue, "^(?:[A-Za-z]:\\|\\\\|/)") Then
		; Абсолютный путь - используем как есть
		$sFullPath = $sInputValue
	Else
		; Относительный путь - ищем относительно папки текущего видеофайла
		Local $sContextDir = ""
		If $nID = $iInput1 And FileExists($sVideoFile1) Then
			$sContextDir = _PathGetDir($sVideoFile1)
		ElseIf $nID = $iInput2 And FileExists($sVideoFile2) Then
			$sContextDir = _PathGetDir($sVideoFile2)
		EndIf

		If $sContextDir = "" Then
			; Нет контекстной папки - сбрасываем поле
			If $nID = $iInput1 Then
				$sVideoFile1 = ""
				GUICtrlSetData($nID, "")
			ElseIf $nID = $iInput2 Then
				$sVideoFile2 = ""
				GUICtrlSetData($nID, "")
			EndIf
			_UpdateFilesInfo()
			Return $GUI_RUNDEFMSG
		EndIf

		$sFullPath = $sContextDir & "\" & $sInputValue
	EndIf

	; Валидация: существование файла и поддерживаемое расширение
	If Not FileExists($sFullPath) Or Not _IsValidVideoExtension($sFullPath) Then
		_SetVideoFile($nID, "")
		Return $GUI_RUNDEFMSG
	EndIf

	; Обновляем данные
	_SetVideoFile($nID, $sFullPath)
	Return $GUI_RUNDEFMSG
EndFunc   ;==>_OnEvent_WM_COMMAND
