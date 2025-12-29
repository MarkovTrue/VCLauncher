#pragma compile(Out, VCLauncher.exe)
#pragma compile(Icon, Assets\icon.ico)

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
Global Const $sAppName = "VCLauncher 0.22"
Global Const $sDirIni = @ScriptDir & '\VCLauncher.ini'
Global Const $sDirFFprobe = @ScriptDir & '\ffprobe.exe'
Global Const $sDirVideoCompare = @ScriptDir & '\video-compare.exe'

Global Const $iGuiWidth = 460
Global Const $iGuiHeight = 120

; Ссылки на элементы GUI
Global $hGui, $iInput1, $iInput2, $iButtonChoose1, $iButtonChoose2, $iButtonCompare, $iRadioDirect, $iRadioVertical
; Остальные переменные
Global $sDirFile1 = IniRead($sDirIni, "LastDirs", "Video1", "")
Global $sDirFile2 = IniRead($sDirIni, "LastDirs", "Video2", "")


If Not FileExists($sDirVideoCompare) Then
	MsgBox(48, $sAppName, 'Не найден "video-compare.exe", а должен лежать тут' & @CR & $sDirFFprobe)
	Exit
EndIf

If Not FileExists($sDirFFprobe) Then
	MsgBox(48, $sAppName, 'Не найден "ffprobe.exe", а должен лежать тут' & @CR & $sDirFFprobe)
	Exit
EndIf

_CreateGUI()
_DefineEvents()

While 1
	Sleep(50)
WEnd

; --- ФУНКЦИИ ---

Func _CreateGUI()
	$hGui = GUICreate($sAppName, $iGuiWidth, $iGuiHeight, -1, -1, $WS_SIZEBOX + $WS_SYSMENU + $WS_MINIMIZEBOX)

	Local $iLabel1 = GUICtrlCreateLabel("Файл 1", 10, 10, 55, 20)
	GUICtrlSetResizing($iLabel1, $GUI_DOCKLEFT + $GUI_DOCKTOP + $GUI_DOCKWIDTH + $GUI_DOCKHEIGHT)

	$iInput1 = GUICtrlCreateInput("", 60, 8, $iGuiWidth - 150, 20)
	GUICtrlSetResizing($iInput1, $GUI_DOCKLEFT + $GUI_DOCKTOP + $GUI_DOCKRIGHT + $GUI_DOCKHEIGHT)
	If $sDirFile1 <> "" Then GUICtrlSetData($iInput1, _GetFileName($sDirFile1))

	$iButtonChoose1 = GUICtrlCreateButton("Выбрать", $iGuiWidth - 82, 7, 74, 22)
	GUICtrlSetResizing($iButtonChoose1, $GUI_DOCKRIGHT + $GUI_DOCKTOP + $GUI_DOCKWIDTH + $GUI_DOCKHEIGHT)

	Local $iLabel2 = GUICtrlCreateLabel("Файл 2", 10, 38, 55, 20)
	GUICtrlSetResizing($iLabel2, $GUI_DOCKLEFT + $GUI_DOCKTOP + $GUI_DOCKWIDTH + $GUI_DOCKHEIGHT)

	$iInput2 = GUICtrlCreateInput("", 60, 36, $iGuiWidth - 150, 20)
	GUICtrlSetResizing($iInput2, $GUI_DOCKLEFT + $GUI_DOCKTOP + $GUI_DOCKRIGHT + $GUI_DOCKHEIGHT)
	If $sDirFile2 <> "" Then GUICtrlSetData($iInput2, _GetFileName($sDirFile2))

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

	GUISetState(@SW_SHOW)
EndFunc   ;==>_CreateGUI

Func _DefineEvents()
	GUICtrlSetOnEvent($iButtonChoose1, "_OnEvent_ButtonChoose1")
	GUICtrlSetOnEvent($iButtonChoose2, "_OnEvent_ButtonChoose2")
	GUICtrlSetOnEvent($iButtonCompare, "_OnEvent_ButtonCompare")
	GUISetOnEvent($GUI_EVENT_CLOSE, "_OnEvent_GUI_EVENT_CLOSE")

	; Ограничение на высоту окна через WM_GETMINMAXINFO
	GUIRegisterMsg($WM_GETMINMAXINFO, "_OnEvent_WM_GETMINMAXINFO")
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

Func _OnEvent_ButtonChoose1()
	Local $sFile1 = FileOpenDialog("Выберите видео 1", _PathGetDir($sDirFile1), "Видео (*.mp4;*.avi;*.mkv;*.mov;*.wmv)|Все файлы (*.*)", 1, _GetFileName($sDirFile1))
	If Not @error And $sFile1 <> "" Then
		$sDirFile1 = $sFile1
		GUICtrlSetData($iInput1, _GetFileName($sDirFile1))
		IniWrite($sDirIni, "LastDirs", "Video1", $sDirFile1)
	EndIf
EndFunc   ;==>_OnEvent_ButtonChoose1

Func _OnEvent_ButtonChoose2()
	Local $sFile2 = FileOpenDialog("Выберите видео 2", _PathGetDir($sDirFile2), "Видео (*.mp4;*.avi;*.mkv;*.mov;*.wmv)|Все файлы (*.*)", 1, _GetFileName($sDirFile2))
	If Not @error And $sFile2 <> "" Then
		$sDirFile2 = $sFile2
		GUICtrlSetData($iInput2, _GetFileName($sDirFile2))
		IniWrite($sDirIni, "LastDirs", "Video2", $sDirFile2)
	EndIf
EndFunc   ;==>_OnEvent_ButtonChoose2


Func _OnEvent_ButtonCompare()
    ; Строки и управляющие ключи
    Local $sCropKeyL = "", $sCropKeyR = "", $sRunKey   = ""

    ; Размер рабочей области экрана
    Local $tDesktopRect   = _WinAPI_GetWorkArea()
    Local $iDesktopWidth  = DllStructGetData($tDesktopRect, "Right")
    Local $iDesktopHeight = DllStructGetData($tDesktopRect, "Bottom")

    ; Получение разрешений видео
    Local $sVideoRes1 = _GetVideoResolution($sDirFile1)
    Local $sVideoRes2 = _GetVideoResolution($sDirFile2)

    Local $iVideo1Width  = _GetWidthFromFFprobeString($sVideoRes1)
    Local $iVideo1Height = _GetHeightFromFFprobeString($sVideoRes1)
    Local $iVideo2Width  = _GetWidthFromFFprobeString($sVideoRes2)
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
                $sCropKeyL = "-l crop=iw:ih-" & $iOrigDiff & " "
                $iCropVideo1Height -= $iOrigDiff
            Else
                ; Crop видео 2
                Local $iOrigDiff = Round($iDiffHeight / $nScale2)
                $sCropKeyR = "-r crop=iw:ih-" & $iOrigDiff & " "
                $iCropVideo2Height -= $iOrigDiff
            EndIf
        EndIf
    EndIf

    ; Формирование команды video-compare
    $sRunKey = '"' & $sDirVideoCompare & '"'

    ; Режим сравнения
    Local $iMaxVideoWidth, $iMaxVideoHeight

    If GUICtrlRead($iRadioVertical) = $GUI_CHECKED Then
        ; Вертикальное сравнение
        $iMaxVideoWidth  = _Max($iVideo1Width, $iVideo2Width)
        $iMaxVideoHeight = _Max($iCropVideo1Height, $iCropVideo2Height) * 2

        If $iMaxVideoWidth > $iDesktopWidth Or $iMaxVideoHeight > $iDesktopHeight Then
            $sRunKey &= " -W"
        EndIf

        $sRunKey &= " -m vstack"
    Else
        ; Горизонтальное сравнение
        $iMaxVideoWidth  = _Max($iVideo1Width, $iVideo2Width)
        $iMaxVideoHeight = _Max($iCropVideo1Height, $iCropVideo2Height)

        If $iMaxVideoWidth > $iDesktopWidth Or $iMaxVideoHeight > $iDesktopHeight Then
            $sRunKey &= " -W"
        EndIf
    EndIf

    ; Добавляем crop и пути к файлам
    $sRunKey &= " " & $sCropKeyL & $sCropKeyR & _
                '"' & $sDirFile1 & '" "' & $sDirFile2 & '"'

    ; Запуск сравнения
    _RunVideoCompare($sRunKey)

EndFunc   ;==> _OnEvent_ButtonCompare




Func _RunVideoCompare($sRunKeys)
	ConsoleWrite($sRunKeys & @CR)

	Local $sOutput = "", $sLine
	; Выполняем
	Local $iProcessPid = Run($sRunKeys)
	; Читаем вывод, как чек возможных ошибок
	While 1
		$sLine = StdoutRead($iProcessPid)
		If @error Then ExitLoop
		$sOutput &= $sLine
	WEnd

	If StringInStr($sOutput, "Error:") Then	MsgBox(48, $sAppName, 'Ошибка' & @CR & $sOutput)

EndFunc


Func _GetVideoResolution($sDirVideo)
	Local $sOutput = "", $sLine
	; Команда ffprobe
	Local $sRunKeys = '"' & $sDirFFprobe & '" -v error -select_streams v:0 -show_entries stream=width,height -of csv=p=0 "' & $sDirVideo & '"'
	; Выполняем
	Local $iProcessPid = Run($sRunKeys, "", @SW_HIDE, $STDERR_CHILD + $STDOUT_CHILD)
	; Читаем вывод
	While 1
		$sLine = StdoutRead($iProcessPid)
		If @error Then ExitLoop
		$sOutput &= $sLine
	WEnd

	ConsoleWrite(@ComSpec & " /c " & $sRunKeys & @CR)
;~ 	MsgBox(48, $sAppName, $sDirVideo & @CR & $sOutput)
	ConsoleWrite($sDirVideo & @CR & $sOutput & @CR)

	Return $sOutput
EndFunc


Func _GetWidthFromFFprobeString($sVar)
	Local $aLineSplit = StringSplit($sVar, ',')
	If $aLineSplit[0] >= 2 Then
		Return Int($aLineSplit[1])
	EndIf
;~ 	MsgBox(48, $sAppName, '_GetWidthFromFFprobeString ошибка' & @CR & $sVar)
	Return 0
EndFunc

Func _GetHeightFromFFprobeString($sVar)
	Local $aLineSplit = StringSplit($sVar, ',')
	If $aLineSplit[0] >= 2 Then
		Return Int($aLineSplit[2])
	EndIf
;~ 	MsgBox(48, $sAppName, '_GetHeightFromFFprobeString ошибка' & @CR & $sVar)
	Return 0
EndFunc


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
