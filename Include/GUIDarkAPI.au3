#include-once
#include <StructureConstants.au3>
#include <AutoItConstants.au3>
#include <WinAPIGdiInternals.au3>

; #INDEX# =======================================================================================================================
; Title .........: GUIDarkAPI UDF Library for AutoIt3
; AutoIt Version : 3.3.18.0
; Language ......: English
; Description ...: API support library for GUIDarkTheme UDF
; Author(s) .....: WildByDesign (including code from NoNameCode, argumentum, UEZ)
; Version .......: 1.6.0.0
; ===============================================================================================================================

; #INTERNAL_USE_ONLY# ===========================================================================================================
; TODO: Need to fill list of all functions.
; ===============================================================================================================================

; #GLOBAL VARIABLES# ============================================================================================================
Global $__DM_g_hDllGdi32 = DllOpen("gdi32.dll")
Global $__DM_g_hDllUser32 = DllOpen("user32.dll")
Global $__DM_g_hDllKernel32 = DllOpen("kernel32.dll")
Global $__DM_g_hDllShlwapi = DllOpen("shlwapi.dll")
Global $__DM_g_hDllUxtheme = DllOpen("uxtheme.dll")
Global $__DM_g_hDllComctl32 = DllOpen("comctl32.dll")
; ===============================================================================================================================

; #GLOBAL CONSTANTS# ============================================================================================================
Global Const $APPMODE_DEFAULT = 0
Global Const $APPMODE_ALLOWDARK = 1
Global Const $APPMODE_FORCEDARK = 2
Global Const $APPMODE_FORCELIGHT = 3
Global Const $APPMODE_MAX = 4
; ===============================================================================================================================

OnAutoItExitRegister("_UnloadDLLs")

Func _UnloadDLLs()
	If $__DM_g_hDllGdi32 Then DllClose($__DM_g_hDllGdi32)
	If $__DM_g_hDllUser32 Then DllClose($__DM_g_hDllUser32)
	If $__DM_g_hDllKernel32 Then DllClose($__DM_g_hDllKernel32)
	If $__DM_g_hDllShlwapi Then DllClose($__DM_g_hDllShlwapi)
	If $__DM_g_hDllUxtheme Then DllClose($__DM_g_hDllUxtheme)
	If $__DM_g_hDllComctl32 Then DllClose($__DM_g_hDllComctl32)
EndFunc   ;==>_UnloadDLLs

Func __WinAPI_GetDpiForWindow($hWnd)
	Local $aResult = DllCall($__DM_g_hDllUser32, "uint", "GetDpiForWindow", "hwnd", $hWnd) ;requires Win10 v1607+ / no server support
	If Not IsArray($aResult) Or @error Then Return SetError(1, @extended, 0)
	If Not $aResult[0] Then Return SetError(2, @extended, 0)
	Return $aResult[0]
EndFunc   ;==>__WinAPI_GetDpiForWindow

Func __WinAPI_GetWindowRect($hWnd)
	Local $tRECT = DllStructCreate($tagRECT)
	Local $aCall = DllCall($__DM_g_hDllUser32, "bool", "GetWindowRect", "hwnd", $hWnd, "struct*", $tRECT)
	If @error Or Not $aCall[0] Then Return SetError(@error + 10, @extended, 0)

	Return $tRECT
EndFunc   ;==>__WinAPI_GetWindowRect

Func __WinAPI_IsWindowVisible($hWnd)
	Local $aCall = DllCall($__DM_g_hDllUser32, "bool", "IsWindowVisible", "hwnd", $hWnd)
	If @error Then Return SetError(@error, @extended, 0)

	Return $aCall[0]
EndFunc   ;==>__WinAPI_IsWindowVisible

Func __WinAPI_GetWindow($hWnd, $iCmd)
	Local $aCall = DllCall($__DM_g_hDllUser32, "hwnd", "GetWindow", "hwnd", $hWnd, "uint", $iCmd)
	If @error Then Return SetError(@error, @extended, 0)

	Return $aCall[0]
EndFunc   ;==>__WinAPI_GetWindow

Func __WinAPI_CallWindowProc($pPrevWndFunc, $hWnd, $iMsg, $wParam, $lParam)
	Local $aCall = DllCall($__DM_g_hDllUser32, "lresult", "CallWindowProc", "ptr", $pPrevWndFunc, "hwnd", $hWnd, "uint", $iMsg, _
			"wparam", $wParam, "lparam", $lParam)
	If @error Then Return SetError(@error, @extended, -1)

	Return $aCall[0]
EndFunc   ;==>__WinAPI_CallWindowProc

Func __WinAPI_EndPaint($hWnd, ByRef $tPAINTSTRUCT)
	Local $aCall = DllCall($__DM_g_hDllUser32, 'bool', 'EndPaint', 'hwnd', $hWnd, 'struct*', $tPAINTSTRUCT)
	If @error Then Return SetError(@error, @extended, False)

	Return $aCall[0]
EndFunc   ;==>__WinAPI_EndPaint

Func __WinAPI_RedrawWindow($hWnd, $tRECT = 0, $hRegion = 0, $iFlags = 5)
	Local $aCall = DllCall($__DM_g_hDllUser32, "bool", "RedrawWindow", "hwnd", $hWnd, "struct*", $tRECT, "handle", $hRegion, _
			"uint", $iFlags)
	If @error Then Return SetError(@error, @extended, False)

	Return $aCall[0]
EndFunc   ;==>__WinAPI_RedrawWindow

Func __WinAPI_GetSystemMetrics($iIndex)
	Local $aCall = DllCall($__DM_g_hDllUser32, "int", "GetSystemMetrics", "int", $iIndex)
	If @error Then Return SetError(@error, @extended, 0)

	Return $aCall[0]
EndFunc   ;==>__WinAPI_GetSystemMetrics

Func __WinAPI_GetClientRect($hWnd)
	Local $tRECT = DllStructCreate($tagRECT)
	Local $aCall = DllCall($__DM_g_hDllUser32, "bool", "GetClientRect", "hwnd", $hWnd, "struct*", $tRECT)
	If @error Or Not $aCall[0] Then Return SetError(@error + 10, @extended, 0)

	Return $tRECT
EndFunc   ;==>__WinAPI_GetClientRect

Func __WinAPI_OffsetRect(ByRef $tRECT, $iDX, $iDY)
	Local $aCall = DllCall($__DM_g_hDllUser32, 'bool', 'OffsetRect', 'struct*', $tRECT, 'int', $iDX, 'int', $iDY)
	If @error Then Return SetError(@error, @extended, 0)
	; If Not $aCall[0] Then Return SetError(1000, 0, 0)

	Return $aCall[0]
EndFunc   ;==>__WinAPI_OffsetRect

Func __WinAPI_GetDCEx($hWnd, $hRgn, $iFlags)
	Local $aCall = DllCall($__DM_g_hDllUser32, 'handle', 'GetDCEx', 'hwnd', $hWnd, 'handle', $hRgn, 'dword', $iFlags)
	If @error Then Return SetError(@error, @extended, 0)
	; If Not $aCall[0] Then Return SetError(1000, 0, 0)

	Return $aCall[0]
EndFunc   ;==>__WinAPI_GetDCEx

Func __WinAPI_FillRect($hDC, $tRECT, $hBrush)
	Local $aCall
	If IsPtr($hBrush) Then
		$aCall = DllCall($__DM_g_hDllUser32, "int", "FillRect", "handle", $hDC, "struct*", $tRECT, "handle", $hBrush)
	Else
		$aCall = DllCall($__DM_g_hDllUser32, "int", "FillRect", "handle", $hDC, "struct*", $tRECT, "dword_ptr", $hBrush)
	EndIf
	If @error Then Return SetError(@error, @extended, False)

	Return $aCall[0]
EndFunc   ;==>__WinAPI_FillRect

Func __WinAPI_ReleaseDC($hWnd, $hDC)
	Local $aCall = DllCall($__DM_g_hDllUser32, "int", "ReleaseDC", "hwnd", $hWnd, "handle", $hDC)
	If @error Then Return SetError(@error, @extended, False)

	Return $aCall[0]
EndFunc   ;==>__WinAPI_ReleaseDC

Func __WinAPI_LockWindowUpdate($hWnd)
	Local $aCall = DllCall($__DM_g_hDllUser32, 'bool', 'LockWindowUpdate', 'hwnd', $hWnd)
	If @error Then Return SetError(@error, @extended, False)
	; If Not $aCall[0] Then Return SetError(1000, 0, 0)

	Return $aCall[0]
EndFunc   ;==>__WinAPI_LockWindowUpdate

Func __WinAPI_DefWindowProc($hWnd, $iMsg, $wParam, $lParam)
	Local $aCall = DllCall($__DM_g_hDllUser32, "lresult", "DefWindowProc", "hwnd", $hWnd, "uint", $iMsg, "wparam", $wParam, _
			"lparam", $lParam)
	If @error Then Return SetError(@error, @extended, 0)

	Return $aCall[0]
EndFunc   ;==>__WinAPI_DefWindowProc

Func __WinAPI_GetWindowDC($hWnd)
	Local $aCall = DllCall($__DM_g_hDllUser32, "handle", "GetWindowDC", "hwnd", $hWnd)
	If @error Then Return SetError(@error, @extended, 0)

	Return $aCall[0]
EndFunc   ;==>__WinAPI_GetWindowDC

Func __WinAPI_GetParent($hWnd)
	Local $aCall = DllCall($__DM_g_hDllUser32, "hwnd", "GetParent", "hwnd", $hWnd)
	If @error Then Return SetError(@error, @extended, 0)

	Return $aCall[0]
EndFunc   ;==>__WinAPI_GetParent

Func __WinAPI_GetDC($hWnd)
	Local $aCall = DllCall($__DM_g_hDllUser32, "handle", "GetDC", "hwnd", $hWnd)
	If @error Then Return SetError(@error, @extended, 0)

	Return $aCall[0]
EndFunc   ;==>__WinAPI_GetDC

Func __WinAPI_BeginPaint($hWnd, ByRef $tPAINTSTRUCT)
	Local Const $tagPAINTSTRUCT = 'hwnd hDC;int fErase;dword rPaint[4];int fRestore;int fIncUpdate;byte rgbReserved[32]'
	$tPAINTSTRUCT = DllStructCreate($tagPAINTSTRUCT)
	Local $aCall = DllCall($__DM_g_hDllUser32, 'handle', 'BeginPaint', 'hwnd', $hWnd, 'struct*', $tPAINTSTRUCT)
	If @error Then Return SetError(@error, @extended, 0)
	; If Not $aCall[0] Then Return SetError(1000, 0, 0)

	Return $aCall[0]
EndFunc   ;==>__WinAPI_BeginPaint

Func __WinAPI_GetClassName($hWnd)
	If Not IsHWnd($hWnd) Then $hWnd = GUICtrlGetHandle($hWnd)
	Local $aCall = DllCall($__DM_g_hDllUser32, "int", "GetClassNameW", "hwnd", $hWnd, "wstr", "", "int", 4096)
	If @error Or Not $aCall[0] Then Return SetError(@error, @extended, '')

	Return SetExtended($aCall[0], $aCall[2])
EndFunc   ;==>__WinAPI_GetClassName

Func __WinAPI_SetActiveWindow($hWnd)
	Local $aCall = DllCall($__DM_g_hDllUser32, 'int', 'SetActiveWindow', 'hwnd', $hWnd)
	If @error Then Return SetError(@error, @extended, 0)
	; If Not $aCall[0] Then Return SetError(1000, 0, 0)

	Return $aCall[0]
EndFunc   ;==>__WinAPI_SetActiveWindow

Func __WinAPI_CallNextHookEx($hHook, $iCode, $wParam, $lParam)
	Local $aCall = DllCall($__DM_g_hDllUser32, "lresult", "CallNextHookEx", "handle", $hHook, "int", $iCode, "wparam", $wParam, "lparam", $lParam)
	If @error Then Return SetError(@error, @extended, -1)

	Return $aCall[0]
EndFunc   ;==>__WinAPI_CallNextHookEx

Func __WinAPI_SetWindowsHookEx($iHook, $pProc, $hDll, $iThreadId = 0)
	Local $aCall = DllCall($__DM_g_hDllUser32, "handle", "SetWindowsHookEx", "int", $iHook, "ptr", $pProc, "handle", $hDll, _
			"dword", $iThreadId)
	If @error Then Return SetError(@error, @extended, 0)

	Return $aCall[0]
EndFunc   ;==>__WinAPI_SetWindowsHookEx

Func __WinAPI_ShowWindow($hWnd, $iCmdShow = @SW_SHOW)
	Local $aCall = DllCall($__DM_g_hDllUser32, "bool", "ShowWindow", "hwnd", $hWnd, "int", $iCmdShow)
	If @error Then Return SetError(@error, @extended, False)

	Return $aCall[0]
EndFunc   ;==>__WinAPI_ShowWindow

Func __WinAPI_UnhookWindowsHookEx($hHook)
	Local $aCall = DllCall($__DM_g_hDllUser32, "bool", "UnhookWindowsHookEx", "handle", $hHook)
	If @error Then Return SetError(@error, @extended, False)

	Return $aCall[0]
EndFunc   ;==>__WinAPI_UnhookWindowsHookEx

Func __WinAPI_DrawText($hDC, $sText, ByRef $tRECT, $iFlags)
	Local $aCall = DllCall($__DM_g_hDllUser32, "int", "DrawTextW", "handle", $hDC, "wstr", $sText, "int", -1, "struct*", $tRECT, _
			"uint", $iFlags)
	If @error Then Return SetError(@error, @extended, 0)

	Return $aCall[0]
EndFunc   ;==>__WinAPI_DrawText

Func __WinAPI_SetFocus($hWnd)
	Local $aCall = DllCall($__DM_g_hDllUser32, "hwnd", "SetFocus", "hwnd", $hWnd)
	If @error Then Return SetError(@error, @extended, 0)

	Return $aCall[0]
EndFunc   ;==>__WinAPI_SetFocus

Func __WinAPI_ScreenToClient($hWnd, ByRef $tPoint)
	Local $aCall = DllCall($__DM_g_hDllUser32, "bool", "ScreenToClient", "hwnd", $hWnd, "struct*", $tPoint)
	If @error Then Return SetError(@error, @extended, False)

	Return $aCall[0]
EndFunc   ;==>__WinAPI_ScreenToClient

Func __WinAPI_InvertRect($hDC, ByRef $tRECT)
	Local $aCall = DllCall($__DM_g_hDllUser32, 'bool', 'InvertRect', 'handle', $hDC, 'struct*', $tRECT)
	If @error Then Return SetError(@error, @extended, False)
	; If Not $aCall[0] Then Return SetError(1000, 0, 0)

	Return $aCall[0]
EndFunc   ;==>__WinAPI_InvertRect

Func __WinAPI_InvalidateRect($hWnd, $tRECT = 0, $bErase = True)
	If Not IsHWnd($hWnd) Then $hWnd = GUICtrlGetHandle($hWnd)
	Local $aCall = DllCall($__DM_g_hDllUser32, "bool", "InvalidateRect", "hwnd", $hWnd, "struct*", $tRECT, "bool", $bErase)
	If @error Then Return SetError(@error, @extended, False)

	Return $aCall[0]
EndFunc   ;==>__WinAPI_InvalidateRect

Func __WinAPI_GetFocus()
	Local $aCall = DllCall($__DM_g_hDllUser32, "hwnd", "GetFocus")
	If @error Then Return SetError(@error, @extended, 0)

	Return $aCall[0]
EndFunc   ;==>__WinAPI_GetFocus

Func __WinAPI_GetDlgItem($hWnd, $iItemID)
	Local $aCall = DllCall($__DM_g_hDllUser32, "hwnd", "GetDlgItem", "hwnd", $hWnd, "int", $iItemID)
	If @error Then Return SetError(@error, @extended, 0)

	Return $aCall[0]
EndFunc   ;==>__WinAPI_GetDlgItem

Func __WinAPI_GetWindowLong($hWnd, $iIndex)
	Local $sFuncName = "GetWindowLongW"
	If @AutoItX64 Then $sFuncName = "GetWindowLongPtrW"
	Local $aCall = DllCall($__DM_g_hDllUser32, "long_ptr", $sFuncName, "hwnd", $hWnd, "int", $iIndex)
	If @error Or Not $aCall[0] Then Return SetError(@error + 10, @extended, 0)

	Return $aCall[0]
EndFunc   ;==>__WinAPI_GetWindowLong

Func __WinAPI_SetWindowLong($hWnd, $iIndex, $iValue)
	__WinAPI_SetLastError(0) ; as suggested in MSDN
	Local $sFuncName = "SetWindowLongW"
	If @AutoItX64 Then $sFuncName = "SetWindowLongPtrW"
	Local $aCall = DllCall($__DM_g_hDllUser32, "long_ptr", $sFuncName, "hwnd", $hWnd, "int", $iIndex, "long_ptr", $iValue)
	If @error Then Return SetError(@error, @extended, 0)

	Return $aCall[0]
EndFunc   ;==>__WinAPI_SetWindowLong

Func __WinAPI_SetWindowPos($hWnd, $hAfter, $iX, $iY, $iCX, $iCY, $iFlags)
	Local $aCall = DllCall($__DM_g_hDllUser32, "bool", "SetWindowPos", "hwnd", $hWnd, "hwnd", $hAfter, "int", $iX, "int", $iY, _
			"int", $iCX, "int", $iCY, "uint", $iFlags)
	If @error Then Return SetError(@error, @extended, False)

	Return $aCall[0]
EndFunc   ;==>__WinAPI_SetWindowPos

Func __WinAPI_SetClassLongEx($hWnd, $iIndex, $iNewLong)
	Local $aCall
	If @AutoItX64 Then
		$aCall = DllCall($__DM_g_hDllUser32, 'ulong_ptr', 'SetClassLongPtrW', 'hwnd', $hWnd, 'int', $iIndex, 'long_ptr', $iNewLong)
	Else
		$aCall = DllCall($__DM_g_hDllUser32, 'dword', 'SetClassLongW', 'hwnd', $hWnd, 'int', $iIndex, 'long', $iNewLong)
	EndIf
	If @error Then Return SetError(@error, @extended, 0)
	; If Not $aCall[0] Then Return SetError(1000, 0, 0)

	Return $aCall[0]
EndFunc   ;==>__WinAPI_SetClassLongEx

Func __WinAPI_TrackMouseEvent($hWnd, $iFlags, $iTime = -1)
	Local $tTME = DllStructCreate('dword;dword;hwnd;dword')
	DllStructSetData($tTME, 1, DllStructGetSize($tTME))
	DllStructSetData($tTME, 2, $iFlags)
	DllStructSetData($tTME, 3, $hWnd)
	DllStructSetData($tTME, 4, $iTime)

	Local $aCall = DllCall($__DM_g_hDllUser32, 'bool', 'TrackMouseEvent', 'struct*', $tTME)
	If @error Then Return SetError(@error, @extended, 0)
	; If Not $aCall[0] Then Return SetError(1000, 0, 0)

	Return $aCall[0]
EndFunc   ;==>__WinAPI_TrackMouseEvent

Func __WinAPI_FindWindowEx($hParent, $hAfter, $sClass, $sTitle = "")
	Local $ret = DllCall($__DM_g_hDllUser32, "hwnd", "FindWindowExW", "hwnd", $hParent, "hwnd", $hAfter, "wstr", $sClass, "wstr", $sTitle)
	If @error Or Not IsArray($ret) Then Return 0
	Return $ret[0]
EndFunc   ;==>__WinAPI_FindWindowEx

Func __WinAPI_ClientToScreen($hWnd, ByRef $tPoint)
	Local $aCall = DllCall($__DM_g_hDllUser32, "bool", "ClientToScreen", "hwnd", $hWnd, "struct*", $tPoint)
	If @error Or Not $aCall[0] Then Return SetError(@error + 10, @extended, 0)

	Return $tPoint
EndFunc   ;==>__WinAPI_ClientToScreen

Func __WinAPI_GetForegroundWindow()
	Local $aCall = DllCall($__DM_g_hDllUser32, "hwnd", "GetForegroundWindow")
	If @error Then Return SetError(@error, @extended, 0)

	Return $aCall[0]
EndFunc   ;==>__WinAPI_GetForegroundWindow

Func __WinAPI_EnumProcessWindows($iPID = 0, $bVisible = True)
	Local $aThreads = __WinAPI_EnumProcessThreads($iPID)
	If @error Then Return SetError(@error, @extended, 0)

	Local $hEnumProc = DllCallbackRegister('__EnumWindowsProc', 'bool', 'hwnd;lparam')

	Dim $__g_vEnum[101][2] = [[0]]
	For $i = 1 To $aThreads[0]
		DllCall($__DM_g_hDllUser32, 'bool', 'EnumThreadWindows', 'dword', $aThreads[$i], 'ptr', DllCallbackGetPtr($hEnumProc), _
				'lparam', $bVisible)
		If @error Then
			ExitLoop
		EndIf
	Next
	DllCallbackFree($hEnumProc)
	If Not $__g_vEnum[0][0] Then Return SetError(11, 0, 0)

	___Inc($__g_vEnum, -1)
	Return $__g_vEnum
EndFunc   ;==>__WinAPI_EnumProcessWindows

Func __WinAPI_EnumChildWindows($hWnd, $bVisible = True)
	If Not __WinAPI_GetWindow($hWnd, 5) Then Return SetError(2, 0, 0) ; $GW_CHILD

	Local $hEnumProc = DllCallbackRegister('__EnumWindowsProc', 'bool', 'hwnd;lparam')

	Dim $__g_vEnum[101][2] = [[0]]
	DllCall($__DM_g_hDllUser32, 'bool', 'EnumChildWindows', 'hwnd', $hWnd, 'ptr', DllCallbackGetPtr($hEnumProc), 'lparam', $bVisible)
	If @error Or Not $__g_vEnum[0][0] Then
		$__g_vEnum = @error + 10
	EndIf
	DllCallbackFree($hEnumProc)
	If $__g_vEnum Then Return SetError($__g_vEnum, 0, 0)

	___Inc($__g_vEnum, -1)
	Return $__g_vEnum
EndFunc   ;==>__WinAPI_EnumChildWindows

Func ___Inc(ByRef $aData, $iIncrement = 100)
	Select
		Case UBound($aData, $UBOUND_COLUMNS)
			If $iIncrement < 0 Then
				ReDim $aData[$aData[0][0] + 1][UBound($aData, $UBOUND_COLUMNS)]
			Else
				$aData[0][0] += 1
				If $aData[0][0] > UBound($aData) - 1 Then
					ReDim $aData[$aData[0][0] + $iIncrement][UBound($aData, $UBOUND_COLUMNS)]
				EndIf
			EndIf
		Case UBound($aData, $UBOUND_ROWS)
			If $iIncrement < 0 Then
				ReDim $aData[$aData[0] + 1]
			Else
				$aData[0] += 1
				If $aData[0] > UBound($aData) - 1 Then
					ReDim $aData[$aData[0] + $iIncrement]
				EndIf
			EndIf
		Case Else
			Return 0
	EndSelect
	Return 1
EndFunc   ;==>___Inc

Func __WinAPI_SetLastError($iErrorCode, Const $_iCallerError = @error, Const $_iCallerExtended = @extended)
	DllCall($__DM_g_hDllKernel32, "none", "SetLastError", "dword", $iErrorCode)
	Return SetError($_iCallerError, $_iCallerExtended, Null)
EndFunc   ;==>__WinAPI_SetLastError

Func __WinAPI_GetCurrentThreadId()
	Local $aCall = DllCall($__DM_g_hDllKernel32, "dword", "GetCurrentThreadId")
	If @error Then Return SetError(@error, @extended, 0)

	Return $aCall[0]
EndFunc   ;==>__WinAPI_GetCurrentThreadId

Func __WinAPI_EnumProcessThreads($iPID = 0)
	If Not $iPID Then $iPID = @AutoItPID

	Local Const $TH32CS_SNAPTHREAD = 0x00000004
	Local $hSnapshot = DllCall($__DM_g_hDllKernel32, 'handle', 'CreateToolhelp32Snapshot', 'dword', $TH32CS_SNAPTHREAD, 'dword', 0)
	If @error Or Not $hSnapshot[0] Then Return SetError(@error + 10, @extended, 0)

	Local Const $tagTHREADENTRY32 = 'dword Size;dword Usage;dword ThreadID;dword OwnerProcessID;long BasePri;long DeltaPri;dword Flags'
	Local $tTHREADENTRY32 = DllStructCreate($tagTHREADENTRY32)
	Local $aRet[101] = [0]

	$hSnapshot = $hSnapshot[0]
	DllStructSetData($tTHREADENTRY32, 'Size', DllStructGetSize($tTHREADENTRY32))
	Local $aCall = DllCall($__DM_g_hDllKernel32, 'bool', 'Thread32First', 'handle', $hSnapshot, 'struct*', $tTHREADENTRY32)
	While Not @error And $aCall[0]
		If DllStructGetData($tTHREADENTRY32, 'OwnerProcessID') = $iPID Then
			___Inc($aRet)
			$aRet[$aRet[0]] = DllStructGetData($tTHREADENTRY32, 'ThreadID')
		EndIf
		$aCall = DllCall($__DM_g_hDllKernel32, 'bool', 'Thread32Next', 'handle', $hSnapshot, 'struct*', $tTHREADENTRY32)
	WEnd
	DllCall($__DM_g_hDllKernel32, "bool", "CloseHandle", "handle", $hSnapshot)
	If Not $aRet[0] Then Return SetError(1, 0, 0)

	___Inc($aRet, -1)
	Return $aRet
EndFunc   ;==>__WinAPI_EnumProcessThreads

Func __SendMessage($hWnd, $iMsg, $wParam = 0, $lParam = 0, $iReturn = 0, $wParamType = "wparam", $lParamType = "lparam", $sReturnType = "lresult")
	Local $aCall = DllCall($__DM_g_hDllUser32, $sReturnType, "SendMessageW", "hwnd", $hWnd, "uint", $iMsg, $wParamType, $wParam, $lParamType, $lParam)
	If @error Then Return SetError(@error, @extended, "")
	If $iReturn >= 0 And $iReturn <= 4 Then Return $aCall[$iReturn]
	Return $aCall
EndFunc   ;==>__SendMessage

Func __WinAPI_CreateRectRgn($iLeftRect, $iTopRect, $iRightRect, $iBottomRect)
	Local $aCall = DllCall($__DM_g_hDllGdi32, "handle", "CreateRectRgn", "int", $iLeftRect, "int", $iTopRect, "int", $iRightRect, _
			"int", $iBottomRect)
	If @error Then Return SetError(@error, @extended, 0)

	Return $aCall[0]
EndFunc   ;==>__WinAPI_CreateRectRgn

Func __WinAPI_CreateSolidBrush($iColor)
	Local $aCall = DllCall($__DM_g_hDllGdi32, "handle", "CreateSolidBrush", "INT", $iColor)
	If @error Then Return SetError(@error, @extended, 0)

	Return $aCall[0]
EndFunc   ;==>__WinAPI_CreateSolidBrush

Func __WinAPI_SetBkColor($hDC, $iColor)
	Local $aCall = DllCall($__DM_g_hDllGdi32, "INT", "SetBkColor", "handle", $hDC, "INT", $iColor)
	If @error Then Return SetError(@error, @extended, -1)

	Return $aCall[0]
EndFunc   ;==>__WinAPI_SetBkColor

Func __WinAPI_DeleteObject($hObject)
	Local $aCall = DllCall($__DM_g_hDllGdi32, "bool", "DeleteObject", "handle", $hObject)
	If @error Then Return SetError(@error, @extended, False)

	Return $aCall[0]
EndFunc   ;==>__WinAPI_DeleteObject

Func __WinAPI_CreatePen($iPenStyle, $iWidth, $iColor)
	Local $aCall = DllCall($__DM_g_hDllGdi32, "handle", "CreatePen", "int", $iPenStyle, "int", $iWidth, "INT", $iColor)
	If @error Then Return SetError(@error, @extended, 0)

	Return $aCall[0]
EndFunc   ;==>__WinAPI_CreatePen

Func __WinAPI_MoveTo($hDC, $iX, $iY)
	Local $aCall = DllCall($__DM_g_hDllGdi32, "bool", "MoveToEx", "handle", $hDC, "int", $iX, "int", $iY, "ptr", 0)
	If @error Then Return SetError(@error, @extended, False)

	Return $aCall[0]
EndFunc   ;==>__WinAPI_MoveTo

Func __WinAPI_Rectangle($hDC, $tRECT)
	Local $aCall = DllCall($__DM_g_hDllGdi32, 'bool', 'Rectangle', 'handle', $hDC, 'int', DllStructGetData($tRECT, 1), _
			'int', DllStructGetData($tRECT, 2), 'int', DllStructGetData($tRECT, 3), 'int', DllStructGetData($tRECT, 4))
	If @error Then Return SetError(@error, @extended, 0)
	; If Not $aCall[0] Then Return SetError(1000, 0, 0)

	Return $aCall[0]
EndFunc   ;==>__WinAPI_Rectangle

Func __WinAPI_LineTo($hDC, $iX, $iY)
	Local $aCall = DllCall($__DM_g_hDllGdi32, "bool", "LineTo", "handle", $hDC, "int", $iX, "int", $iY)
	If @error Then Return SetError(@error, @extended, False)

	Return $aCall[0]
EndFunc   ;==>__WinAPI_LineTo

Func __WinAPI_DrawLine($hDC, $iX1, $iY1, $iX2, $iY2)
	__WinAPI_MoveTo($hDC, $iX1, $iY1)
	If @error Then Return SetError(@error, @extended, False)
	__WinAPI_LineTo($hDC, $iX2, $iY2)
	If @error Then Return SetError(@error + 10, @extended, False)
	Return True
EndFunc   ;==>__WinAPI_DrawLine

Func __WinAPI_BitBlt($hDestDC, $iXDest, $iYDest, $iWidth, $iHeight, $hSrcDC, $iXSrc, $iYSrc, $iROP)
	Local $aCall = DllCall($__DM_g_hDllGdi32, "bool", "BitBlt", "handle", $hDestDC, "int", $iXDest, "int", $iYDest, "int", $iWidth, _
			"int", $iHeight, "handle", $hSrcDC, "int", $iXSrc, "int", $iYSrc, "dword", $iROP)
	If @error Then Return SetError(@error, @extended, False)

	Return $aCall[0]
EndFunc   ;==>__WinAPI_BitBlt

Func __WinAPI_DeleteDC($hDC)
	Local $aCall = DllCall($__DM_g_hDllGdi32, "bool", "DeleteDC", "handle", $hDC)
	If @error Then Return SetError(@error, @extended, False)

	Return $aCall[0]
EndFunc   ;==>__WinAPI_DeleteDC

Func __WinAPI_SelectObject($hDC, $hGDIObj)
	Local $aCall = DllCall($__DM_g_hDllGdi32, "handle", "SelectObject", "handle", $hDC, "handle", $hGDIObj)
	If @error Then Return SetError(@error, @extended, False)

	Return $aCall[0]
EndFunc   ;==>__WinAPI_SelectObject

Func __WinAPI_SetBkMode($hDC, $iBkMode)
	Local $aCall = DllCall($__DM_g_hDllGdi32, "int", "SetBkMode", "handle", $hDC, "int", $iBkMode)
	If @error Then Return SetError(@error, @extended, 0)

	Return $aCall[0]
EndFunc   ;==>__WinAPI_SetBkMode

Func __WinAPI_SetTextColor($hDC, $iColor)
	Local $aCall = DllCall($__DM_g_hDllGdi32, "INT", "SetTextColor", "handle", $hDC, "INT", $iColor)
	If @error Then Return SetError(@error, @extended, -1)

	Return $aCall[0]
EndFunc   ;==>__WinAPI_SetTextColor

Func __WinAPI_GetStockObject($iObject)
	Local $aCall = DllCall($__DM_g_hDllGdi32, "handle", "GetStockObject", "int", $iObject)
	If @error Then Return SetError(@error, @extended, 0)

	Return $aCall[0]
EndFunc   ;==>__WinAPI_GetStockObject

Func __WinAPI_CreateFont($iHeight, $iWidth, $iEscape = 0, $iOrientn = 0, $iWeight = $__WINAPICONSTANT_FW_NORMAL, $bItalic = False, $bUnderline = False, $bStrikeout = False, $iCharset = $__WINAPICONSTANT_DEFAULT_CHARSET, $iOutputPrec = $__WINAPICONSTANT_OUT_DEFAULT_PRECIS, $iClipPrec = $__WINAPICONSTANT_CLIP_DEFAULT_PRECIS, $iQuality = $__WINAPICONSTANT_DEFAULT_QUALITY, $iPitch = $__WINAPICONSTANT_FF_DONTCARE, $sFace = 'Arial')
	Local $aCall = DllCall($__DM_g_hDllGdi32, "handle", "CreateFontW", "int", $iHeight, "int", $iWidth, "int", $iEscape, _
			"int", $iOrientn, "int", $iWeight, "dword", $bItalic, "dword", $bUnderline, "dword", $bStrikeout, _
			"dword", $iCharset, "dword", $iOutputPrec, "dword", $iClipPrec, "dword", $iQuality, "dword", $iPitch, "wstr", $sFace)
	If @error Then Return SetError(@error, @extended, 0)

	Return $aCall[0]
EndFunc   ;==>__WinAPI_CreateFont

Func __WinAPI_CreateCompatibleDC($hDC)
	Local $aCall = DllCall($__DM_g_hDllGdi32, "handle", "CreateCompatibleDC", "handle", $hDC)
	If @error Then Return SetError(@error, @extended, 0)

	Return $aCall[0]
EndFunc   ;==>__WinAPI_CreateCompatibleDC

Func __WinAPI_CreateCompatibleBitmap($hDC, $iWidth, $iHeight)
	Local $aCall = DllCall($__DM_g_hDllGdi32, "handle", "CreateCompatibleBitmap", "handle", $hDC, "int", $iWidth, "int", $iHeight)
	If @error Then Return SetError(@error, @extended, 0)

	Return $aCall[0]
EndFunc   ;==>__WinAPI_CreateCompatibleBitmap

Func __WinAPI_GetObject($hObject, $iSize, $pObject)
	Local $aCall = DllCall($__DM_g_hDllGdi32, "int", "GetObjectW", "handle", $hObject, "int", $iSize, "struct*", $pObject)
	If @error Then Return SetError(@error, @extended, 0)

	Return $aCall[0]
EndFunc   ;==>__WinAPI_GetObject

Func __WinAPI_GetDeviceCaps($hDC, $iIndex)
	Local $aCall = DllCall($__DM_g_hDllGdi32, "int", "GetDeviceCaps", "handle", $hDC, "int", $iIndex)
	If @error Then Return SetError(@error, @extended, 0)

	Return $aCall[0]
EndFunc   ;==>__WinAPI_GetDeviceCaps

Func __WinAPI_GetTextExtentPoint32($hDC, $sText)
	Local $tSize = DllStructCreate($tagSIZE)
	Local $iSize = StringLen($sText)
	Local $aCall = DllCall($__DM_g_hDllGdi32, "bool", "GetTextExtentPoint32W", "handle", $hDC, "wstr", $sText, "int", $iSize, "struct*", $tSize)
	If @error Or Not $aCall[0] Then Return SetError(@error + 10, @extended, 0)

	Return $tSize
EndFunc   ;==>__WinAPI_GetTextExtentPoint32

Func __WinAPI_ColorAdjustLuma($iRGB, $iPercent, $bScale = True)
	If $iRGB = -1 Then Return SetError(10, 0, -1)

	If $bScale Then
		$iPercent = Floor($iPercent * 10)
	EndIf

	Local $aCall = DllCall($__DM_g_hDllShlwapi, 'dword', 'ColorAdjustLuma', 'dword', ___RGB($iRGB), 'int', $iPercent, 'bool', $bScale)
	If @error Then Return SetError(@error, @extended, -1)

	Return ___RGB($aCall[0])
EndFunc   ;==>__WinAPI_ColorAdjustLuma

Func ___RGB($iColor)
	Local $__g_iRGBMode = 1
	If $__g_iRGBMode Then
		$iColor = __WinAPI_SwitchColor($iColor)
	EndIf
	Return $iColor
EndFunc   ;==>___RGB

Func __WinAPI_SwitchColor($iColor)
	If $iColor = -1 Then Return $iColor
	Return BitOR(BitAND($iColor, 0x00FF00), BitShift(BitAND($iColor, 0x0000FF), -16), BitShift(BitAND($iColor, 0xFF0000), 16))
EndFunc   ;==>__WinAPI_SwitchColor

Func __WinAPI_SetWindowTheme($hWnd, $sName = Default, $sList = Default)
	If Not IsString($sName) Then $sName = Null
	If Not IsString($sList) Then $sList = Null

	Local $sResult = DllCall($__DM_g_hDllUxtheme, 'long', 'SetWindowTheme', 'hwnd', $hWnd, 'wstr', $sName, 'wstr', $sList)
	If @error Then Return SetError(@error, @extended, 0)
	If $sResult[0] Then Return SetError(10, $sResult[0], 0)

	Return 1
EndFunc   ;==>__WinAPI_SetWindowTheme

Func __WinAPI_GetThemeAppProperties()
	Local $sResult = DllCall($__DM_g_hDllUxtheme, 'dword', 'GetThemeAppProperties')
	If @error Then Return SetError(@error, @extended, 0)

	Return $sResult[0]
EndFunc   ;==>__WinAPI_GetThemeAppProperties

Func __WinAPI_GetThemePartSize($hTheme, $iPartID, $iStateID, $hDC, $tRECT, $iType)
	Local $tSize = DllStructCreate($tagSIZE)
	Local $sResult = DllCall($__DM_g_hDllUxtheme, 'long', 'GetThemePartSize', 'handle', $hTheme, 'handle', $hDC, 'int', $iPartID, _
			'int', $iStateID, 'struct*', $tRECT, 'int', $iType, 'struct*', $tSize)
	If @error Then Return SetError(@error, @extended, 0)
	If $sResult[0] Then Return SetError(10, $sResult[0], 0)

	Return $tSize
EndFunc   ;==>__WinAPI_GetThemePartSize

Func __WinAPI_SetThemeAppProperties($iFlags)
	DllCall($__DM_g_hDllUxtheme, 'none', 'SetThemeAppProperties', 'dword', $iFlags)
	If @error Then Return SetError(@error, @extended, 0)

	Return 1
EndFunc   ;==>__WinAPI_SetThemeAppProperties

Func __WinAPI_IsDarkModeAllowedForWindow($hWnd)
	Local $aResult = DllCall($__DM_g_hDllUxtheme, "bool", 137, "hwnd", $hWnd)
	If @error Then Return SetError(1, 0, False)
	Return ($aResult[0] <> 0)
EndFunc   ;==>__WinAPI_IsDarkModeAllowedForWindow

Func __WinAPI_DrawThemeBackground($hTheme, $iPartID, $iStateID, $hDC, $tRECT, $tCLIP = 0)
	Local $sResult = DllCall($__DM_g_hDllUxtheme, 'long', 'DrawThemeBackground', 'handle', $hTheme, 'handle', $hDC, 'int', $iPartID, _
			'int', $iStateID, 'struct*', $tRECT, 'struct*', $tCLIP)
	If @error Then Return SetError(@error, @extended, 0)
	If $sResult[0] Then Return SetError(10, $sResult[0], 0)

	Return 1
EndFunc   ;==>__WinAPI_DrawThemeBackground

Func __WinAPI_CloseThemeData($hTheme)
	Local $sResult = DllCall($__DM_g_hDllUxtheme, 'long', 'CloseThemeData', 'handle', $hTheme)
	If @error Then Return SetError(@error, @extended, 0)
	If $sResult[0] Then Return SetError(10, $sResult[0], 0)

	Return 1
EndFunc   ;==>__WinAPI_CloseThemeData

Func __WinAPI_OpenThemeData($hWnd, $sClass)
	Local $sResult = DllCall($__DM_g_hDllUxtheme, 'handle', 'OpenThemeData', 'hwnd', $hWnd, 'wstr', $sClass)
	If @error Then Return SetError(@error, @extended, 0)
	; If Not $sResult[0] Then Return SetError(1000, 0, 0)

	Return $sResult[0]
EndFunc   ;==>__WinAPI_OpenThemeData

Func __WinAPI_OpenThemeDataForDpi($hWnd, $sClass, $iDPI)
	Local $sResult = DllCall($__DM_g_hDllUxtheme, 'handle', 'OpenThemeDataForDpi', 'hwnd', $hWnd, 'wstr', $sClass, "uint", $iDPI)
	If @error Then Return SetError(@error, @extended, 0)
	; If Not $sResult[0] Then Return SetError(1000, 0, 0)

	Return $sResult[0]
EndFunc   ;==>__WinAPI_OpenThemeDataForDpi

Func __WinAPI_AllowDarkModeForWindow($hWnd, $bAllow = True)
	If @OSBuild < 17763 Then Return SetError(-1, 0, False)
	Local $fnAllowDarkModeForWindow = 133
	Local $aResult = DllCall($__DM_g_hDllUxtheme, 'bool', $fnAllowDarkModeForWindow, 'hwnd', $hWnd, 'bool', $bAllow)
	If @error Then Return SetError(@error, @extended, False)
	Return $aResult[0]
EndFunc   ;==>__WinAPI_AllowDarkModeForWindow

Func __WinAPI_FlushMenuThemes()
	If @OSBuild < 17763 Then Return SetError(-1, 0, False)
	Local $fnFlushMenuThemes = 136
	DllCall($__DM_g_hDllUxtheme, 'none', $fnFlushMenuThemes)
	If @error Then Return SetError(@error, @extended, False)
	Return True
EndFunc   ;==>__WinAPI_FlushMenuThemes

Func __WinAPI_RefreshImmersiveColorPolicyState()
	If @OSBuild < 17763 Then Return SetError(-1, 0, False)
	Local $fnRefreshImmersiveColorPolicyState = 104
	DllCall($__DM_g_hDllUxtheme, 'none', $fnRefreshImmersiveColorPolicyState)
	If @error Then Return SetError(@error, @extended, False)
	Return True
EndFunc   ;==>__WinAPI_RefreshImmersiveColorPolicyState

Func __WinAPI_SetPreferredAppMode($PREFERREDAPPMODE)
	If @OSBuild < 18362 Then Return SetError(-1, 0, False)
	Local $fnSetPreferredAppMode = 135
	Local $aResult = DllCall($__DM_g_hDllUxtheme, 'int', $fnSetPreferredAppMode, 'int', $PREFERREDAPPMODE)
	If @error Then Return SetError(@error, @extended, '')
	Return $aResult[0]
EndFunc   ;==>__WinAPI_SetPreferredAppMode

Func __WinAPI_ShouldAppsUseDarkMode()
	Local $aResult = DllCall($__DM_g_hDllUxtheme, "bool", 132)
	If @error Then Return SetError(1, 0, False)
	Return $aResult[0]
EndFunc   ;==>__WinAPI_ShouldAppsUseDarkMode

Func __WinAPI_DwmSetWindowAttribute($hWnd, $iAttribute, $iData)
	Local $aCall = DllCall('dwmapi.dll', 'long', 'DwmSetWindowAttribute', 'hwnd', $hWnd, 'dword', $iAttribute, _
			'dword*', $iData, 'dword', 4)
	If @error Then Return SetError(@error, @extended, 0)
	If $aCall[0] Then Return SetError(10, $aCall[0], 0)
	Return 1
EndFunc   ;==>__WinAPI_DwmSetWindowAttribute

Func __WinAPI_DwmExtendFrameIntoClientArea($hWnd, $tMARGINS = 0)
	If Not IsDllStruct($tMARGINS) Then
		$tMARGINS = _WinAPI_CreateMargins(-1, -1, -1, -1)
	EndIf

	Local $aCall = DllCall('dwmapi.dll', 'long', 'DwmExtendFrameIntoClientArea', 'hwnd', $hWnd, 'struct*', $tMARGINS)
	If @error Then Return SetError(@error, @extended, 0)
	If $aCall[0] Then Return SetError(10, $aCall[0], 0)

	Return 1
EndFunc   ;==>__WinAPI_DwmExtendFrameIntoClientArea

Func __WinAPI_CreateRectEx($iX, $iY, $iWidth, $iHeight)
	Local $tRECT = DllStructCreate($tagRECT)
	DllStructSetData($tRECT, 1, $iX)
	DllStructSetData($tRECT, 2, $iY)
	DllStructSetData($tRECT, 3, $iX + $iWidth)
	DllStructSetData($tRECT, 4, $iY + $iHeight)

	Return $tRECT
EndFunc   ;==>__WinAPI_CreateRectEx

Func __WinAPI_Base64Decode($sB64String)
	Local $aCrypt = DllCall("Crypt32.dll", "bool", "CryptStringToBinaryA", "str", $sB64String, "dword", 0, "dword", 1, "ptr", 0, "dword*", 0, "ptr", 0, "ptr", 0)
	If @error Or Not $aCrypt[0] Then Return SetError(1, 0, "")
	Local $bBuffer = DllStructCreate("byte[" & $aCrypt[5] & "]")
	$aCrypt = DllCall("Crypt32.dll", "bool", "CryptStringToBinaryA", "str", $sB64String, "dword", 0, "dword", 1, "struct*", $bBuffer, "dword*", $aCrypt[5], "ptr", 0, "ptr", 0)
	If @error Or Not $aCrypt[0] Then Return SetError(2, 0, "")
	Return DllStructGetData($bBuffer, 1)
EndFunc   ;==>__WinAPI_Base64Decode

Func _ColorToCOLORREF($iColor) ;RGB to BGR
	Local $iR = BitAND(BitShift($iColor, 16), 0xFF)
	Local $iG = BitAND(BitShift($iColor, 8), 0xFF)
	Local $iB = BitAND($iColor, 0xFF)
	Return BitOR(BitShift($iB, -16), BitShift($iG, -8), $iR)
EndFunc   ;==>_ColorToCOLORREF

Func _GUICtrlTreeView_SetExtendedStyle($hTreeView, $iExStyle)
	Local Const $TVM_SETEXTENDEDSTYLE = 0x112C
	Local $iResult = __SendMessage($hTreeView, $TVM_SETEXTENDEDSTYLE, 0x07FD, $iExStyle)
	Return SetError(@error, @extended, $iResult)
EndFunc   ;==>_GUICtrlTreeView_SetExtendedStyle

Func __WinAPI_DefSubclassProc($hWnd, $iMsg, $wParam, $lParam)
	Local $aCall = DllCall($__DM_g_hDllComctl32, 'lresult', 'DefSubclassProc', 'hwnd', $hWnd, 'uint', $iMsg, 'wparam', $wParam, _
			'lparam', $lParam)
	If @error Then Return SetError(@error, @extended, 0)

	Return $aCall[0]
EndFunc   ;==>__WinAPI_DefSubclassProc

Func __WinAPI_RemoveWindowSubclass($hWnd, $pSubclassProc, $idSubClass)
	Local $aCall = DllCall($__DM_g_hDllComctl32, 'bool', 'RemoveWindowSubclass', 'hwnd', $hWnd, 'ptr', $pSubclassProc, 'uint_ptr', $idSubClass)
	If @error Then Return SetError(@error, @extended, False)
	; If Not $aCall[0] Then Return SetError(1000, 0, 0)

	Return $aCall[0]
EndFunc   ;==>__WinAPI_RemoveWindowSubclass

Func __WinAPI_SetWindowSubclass($hWnd, $pSubclassProc, $idSubClass, $pData = 0)
	Local $aCall = DllCall($__DM_g_hDllComctl32, 'bool', 'SetWindowSubclass', 'hwnd', $hWnd, 'ptr', $pSubclassProc, 'uint_ptr', $idSubClass, _
			'dword_ptr', $pData)
	If @error Then Return SetError(@error, @extended, 0)
	; If Not $aCall[0] Then Return SetError(1000, 0, 0)

	Return $aCall[0]
EndFunc   ;==>__WinAPI_SetWindowSubclass

Func __WinAPI_GetMousePos($bToClient = False, $hWnd = 0)
	Local $iMode = Opt("MouseCoordMode", 1)
	Local $aPos = MouseGetPos()
	Opt("MouseCoordMode", $iMode)

	Local $tPoint = DllStructCreate($tagPOINT)
	DllStructSetData($tPoint, "X", $aPos[0])
	DllStructSetData($tPoint, "Y", $aPos[1])
	If $bToClient And Not __WinAPI_ScreenToClient($hWnd, $tPoint) Then Return SetError(@error + 20, @extended, 0)

	Return $tPoint
EndFunc   ;==>__WinAPI_GetMousePos
