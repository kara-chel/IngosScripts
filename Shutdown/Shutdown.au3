Opt("TrayIconHide", 0) ;0 - отображать, 1 - скрыть
Opt("TrayMenuMode", 3)
Opt("TrayAutoPause", 0)
Opt("TrayOnEventMode", 1)

#include <File.au3>

Const $cVersion = "ShutdownIngos20160406"
Const $cProduct = "Shutdown"
Const $DEBUBMODE = False
Const $cSec = 60; Повторять _MainProcess() каждые $cSec секунд

Global $hLogFile
Global $sLOGFile = @ScriptDir & "\..\Log\" & $cProduct & "\[" & @UserName & "].[" & @ComputerName & "].[" & @IPAddress1 & "].log"
Global $sINIFile = @ScriptDir & "\" & $cProduct & ".ini"
Global $tTimer, $tSleep
Global $aExceptionUserName

Func MainExit()
	_FileWriteLog($hLogFile, "Скрипт завершен")
	FileClose($hLogFile); Закрываем LOG-файл
	Exit
EndFunc

Func MainInit()
	; Проверка на повторный запуск скрипта
	If WinExists($cVersion) Then Exit
	AutoItWinSetTitle($cVersion)
	;$hExist = GUICreate($cVersion, 0, 0, 0, 0)
	;GUISetState(@SW_HIDE)
	; /

	; Таймер
	$tTimer = TimerInit()
	$tSleep = 0
	; /

	$hLogFile = FileOpen($sLogFile, $FO_APPEND)
	If $hLogFile = -1 Then
		MsgBox($MB_SYSTEMMODAL, $cVersion, "Ошибка открытия LOG-файла")
	EndIf
	_FileWriteLog($hLogFile, "Скрипт запущен")
	TrayTip($cProduct, "Скрипт запущен", 3, 1)

	If $DEBUBMODE Then _FileWriteLog($hLogFile, "MainInit(): @ScriptDir = " & @ScriptDir)

	$aExceptionUserName = StringSplit(StringUpper(StringStripWS(IniRead($sINIFile, "Exception", "UserName", ""), 3)), ',')
	If _ArraySearch($aExceptionUserName, StringUpper(@UserName)) >= 0 Then MainExit()

	;TrayMenu
	Local $exititem = TrayCreateItem("Выйти")
	TrayItemSetOnEvent(-1, "MainExit")
	TraySetState()
	TraySetIcon("Shell32.dll", -21)
	;End TrayMenu

EndFunc

Func MainLoop()
	While 1
		if TimerDiff($tTimer) < $tSleep Then
			sleep(100)
		Else
			_MainProcess()
			$tSleep = TimerDiff($tTimer) + $cSec * 1000
		EndIf
	WEnd
EndFunc

Func _MainProcess()
	_FileWriteLog($hLogFile, "_MainProcess(): Shutdown")
	Sleep(500)
	;Shutdown(300)
	Shutdown(1)
EndFunc


;Main
MainInit()
MainLoop()
MainExit()
;End Main
