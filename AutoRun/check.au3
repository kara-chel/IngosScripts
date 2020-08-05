Opt("TrayIconHide", 1)          ;0 - отображать, 1 - скрыть

#include <File.au3>

Const $cVersion = "CheckIngos20160318"
Const $cProduct = "Check"

Global $hLogFile
Global $sLOGFile = @ScriptDir & "\..\Log\" & $cProduct & "\[" & @UserName & "].[" & @ComputerName & "].[" & @IPAddress1 & "].log"

Func MainInit()
	; Проверка на повторный запуск скрипта
	If WinExists($cVersion) Then Exit
	$hExist = GUICreate($cVersion, 0, 0, 0, 0)
	GUISetState(@SW_HIDE)
	; /

	$hLogFile = FileOpen($sLogFile, $FO_APPEND)
	If $hLogFile = -1 Then
		MsgBox($MB_SYSTEMMODAL, $cVersion, "Ошибка открытия LOG-файла.")
	EndIf
	_FileWriteLog($hLogFile, "Скрипт запущен")

	_FileWriteLog($hLogFile, "@LogonDomain: " & @LogonDomain)
	_FileWriteLog($hLogFile, "@ComputerName: " & @ComputerName)
	_FileWriteLog($hLogFile, "@UserName: " & @UserName)

EndFunc

Func MainLoop()
	; :)
EndFunc

Func MainExit()
	_FileWriteLog($hLogFile, "Скрипт завершен")
	FileClose($hLogFile); Закрываем LOF-файл
	Exit
EndFunc

;Main
MainInit()
MainLoop()
MainExit()
;End Main
