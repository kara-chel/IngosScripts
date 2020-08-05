;TODO добавить событие по таймеру
;TODO добавить LOG (просмотр картинок, переходы)
Opt("TrayIconHide", 1) ;0 - отображать, 1 - скрыть

#Include <GUIConstants.au3>
#include <GuiButton.au3>
#include <GDIPlus.au3>
#include <Array.au3>
#include <File.au3>

Const $cVersion = "CongratulateIngos20160321"
Const $cProduct = "Congratulate"

Global $sINIFile = @ScriptDir & '\Congratulate.ini'
Global $hGUI, $hBtnClose, $hBtnNext;, $hGraphic
Global $tTimer, $tNext, $tBtnTimer, $tBtnSleep, $tRefreshTimer, $tRefreshSleep

Global $sFirstPath
Global $sContinuePath
Global $iFirstTime
Global $iNextTime
Global $aExclusion

Global $hGraphic
Global $hBitmap
Global $iGfx_Save
Global $g_hBmp_Buffer
Global $g_hGfx_Buffer
Global $bNCACTIVATE = True

Func _GetRandomFiles($sPath)
	$aFilesList = _FileListToArray($sPath)
	Return $sPath & '\' & $aFilesList[Random(1, $aFilesList[0])]
EndFunc

Func MainInit()

	; Проверка на повторный запуск скрипта
	If WinExists($cVersion) Then Exit
	$hExist = GUICreate($cVersion, 0, 0, 0, 0)
	GUISetState(@SW_HIDE)
	; /

	; Реакция на изменение фокуса
	GUIRegisterMsg($WM_NCACTIVATE, 'WM_NCACTIVATE')

	; Загружаем параметры с INI файла
	$sFirstPath = IniRead($sINIFile, @UserName, "FirstPath", "")
	If $sFirstPath = "" Then $sFirstPath = IniRead($sINIFile, "General", "FirstPath", "img\first")
	$sFirstPath = @ScriptDir & '\' & $sFirstPath

	$sContinuePath = IniRead($sINIFile, @UserName, "ContinuePath", "")
	If $sContinuePath = "" Then $sContinuePath = IniRead($sINIFile, "General", "ContinuePath", "img\900x1600")
	$sContinuePath = @ScriptDir & '\' & $sContinuePath

	Local $sFirstTime = IniRead($sINIFile, @UserName, "FirstTime", "")
	If $sFirstTime = "" Then $sFirstTime = IniRead($sINIFile, "General", "FirstTime", "60")
	$iFirstTime = Int($sFirstTime)

	Local $sNextTime = IniRead($sINIFile, @UserName, "NextTime", "")
	If $sNextTime = "" Then $sNextTime = IniRead($sINIFile, "General", "NextTime", "7")
	$iNextTime = Int($sNextTime)

	Local $sExclusion = IniRead($sINIFile, @UserName, "Exclusion", "")
	If $sExclusion = "" Then $sExclusion = IniRead($sINIFile, "General", "Exclusion", "")
	$aExclusion = StringSplit(StringUpper($sExclusion), ',')
	; /

	MainCheck()

	$tTimer = TimerInit()
	$tNext = 1000 * $iFirstTime
	$tBtnTimer = TimerInit()
	$tBtnSleep = 100
	$tRefreshTimer = TimerInit()
	$tRefreshSleep = 1000

	_GDIPlus_Startup()

	$hGUI = GUICreate($cProduct, 0, 0, -1, -1, $WS_POPUP, $WS_EX_TOPMOST) ;$WS_POPUP
	GUISetState(@SW_MAXIMIZE, $hGUI)

	$hBtnClose = GUICtrlCreateButton("Закрыть", @DesktopWidth - 85 - 2, @DesktopHeight - 25 - 2, 85, 25)
	$hBtnNext = GUICtrlCreateButton("Следующая", @DesktopWidth - 85 - 2 - 95 - 2, @DesktopHeight - 25 - 2, 95, 25)
	GUISetState()

	$hGraphic = _GDIPlus_GraphicsCreateFromHWND($hGUI)

	_ShowImg(_GetRandomFiles($sFirstPath))
EndFunc

Func _ShowImg($sFile)
	;$hGraphic = _GDIPlus_GraphicsCreateFromHWND($hGUI)
	;$hBitmap = _GDIPlus_BitmapCreateFromFile($sFile)

	$hBitmap = _GDIPlus_ImageLoadFromFile($sFile)
	$hBitmap = _GDIPlus_ImageResize($hBitmap, @DesktopWidth, @DesktopHeight)

	;_GDIPlus_BitmapDispose($hBitmap)
	;_GDIPlus_GraphicsDispose($hGraphic)

	_Draw()
EndFunc

Func _Draw()
	_GDIPlus_GraphicsDrawImage($hGraphic, $hBitmap, 0, 0)

	;Фикс - показать кнопку, т.к. картинка еx скрывает
	_GUICtrlButton_Show($hBtnClose, False)
	_GUICtrlButton_Show($hBtnClose, True)
	_GUICtrlButton_Show($hBtnNext, False)
	_GUICtrlButton_Show($hBtnNext, True)
	_GUICtrlButton_SetFocus($hBtnNext, True)
EndFunc

Func MainLoop()
	While 1
		Switch GUIGetMsg()
			Case $GUI_EVENT_CLOSE, $hBtnClose
				ExitLoop
			Case $hBtnNext
				_ShowImg(_GetRandomFiles($sContinuePath))
				$tNext = TimerDiff($tTimer) + 1000 * $iNextTime
		EndSwitch

		if TimerDiff($tTimer) >= $tNext Then
			_ShowImg(_GetRandomFiles($sContinuePath))
			$tNext = TimerDiff($tTimer) + 1000 * $iNextTime
		EndIf
		if TimerDiff($tBtnTimer) >= $tBtnSleep Then
			GUICtrlSetData($hBtnNext, 'Следующая (' & Round(($tNext - TimerDiff($tTimer))/1000) & ')')
			$tBtnSleep = TimerDiff($tBtnTimer) + 100
		EndIf
		if TimerDiff($tRefreshTimer) >= $tRefreshSleep Then
			_Draw()
			$tRefreshSleep = TimerDiff($tBtnTimer) + 1000
		EndIf
	WEnd
EndFunc

Func MainExit()
	_GDIPlus_BitmapDispose($hBitmap)
	_GDIPlus_GraphicsDispose($hGraphic)
    _GDIPlus_Shutdown()
	GUIDelete($hGUI)
	Exit
EndFunc

Func MainCheck()
	Local $aExclusion
	Local $iFind

	;$sExclusion
	$iFind = _ArraySearch($aExclusion, StringUpper(@UserName))
	if  $iFind >= 0 Then
		Exit
	EndIf

	;Проверить OS
	If @OSVersion = "Win_XP" Then
		Exit
	EndIf

EndFunc

Func WM_NCACTIVATE($hWnd, $Msg, $wParam, $lParam)
	If Int(Hex($wParam)) = 0 Then WinActivate($hGUI)

	If Not $bNCACTIVATE And Int(Hex($wParam)) = 1 Then
		;Sleep(1000)
		;_Draw()
		$bNCACTIVATE = True
	ElseIf $bNCACTIVATE And Int(Hex($wParam)) = 0 Then
		$bNCACTIVATE = False
	EndIf
	;ToolTip('Активность: ' & Int(Hex($wParam)) & @CRLF & _
	;	'$hWnd: ' & Int(Hex($hWnd)) & @CRLF & _
	;	'$Msg: ' & Int(Hex($Msg)) & @CRLF & _
	;	'$lParam: ' & Int(Hex($lParam)), 10, 70, 'Главное окно')
EndFunc

MainInit()
MainLoop()
MainExit()
