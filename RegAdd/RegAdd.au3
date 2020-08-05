Opt("TrayIconHide", 0) ;0 - отображать, 1 - скрыть
Opt("TrayMenuMode", 3)
Opt("TrayAutoPause", 0)
Opt("TrayOnEventMode", 1)

#include <File.au3>

Const $cVersion = "TLSIngos20200729"
Const $cProduct = "TLS"
Const $DEBUBMODE = False

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
	$hExist = GUICreate($cVersion, 0, 0, 0, 0)
	GUISetState(@SW_HIDE)
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
	If _ArraySearch($aExceptionUserName, StringUpper(@UserName)) >= 0 Then
		_FileWriteLog($hLogFile, "MainInit(): Expection = " & StringUpper(@UserName))
		MainExit()
	EndIf

	;TrayMenu
	Local $exititem = TrayCreateItem("Выйти")
	TrayItemSetOnEvent(-1, "MainExit")
	TraySetState()
	TraySetIcon("Shell32.dll", -21)
	;End TrayMenu

EndFunc

Func MainLoop()
	_FileWriteLog($hLogFile, "MainLoop()")
	; Создаёт раздел
	$result = RegWrite("HKEY_LOCAL_MACHINE\XXXXXX\XXXXXXXXXXXXXXXXX\XXXXXXX\XXXXXXXXXXXXXXXXX\XXXXXXXX\XXXXXXX")
	If $result = 0 Then
		_FileWriteLog($hLogFile, "MainLoop(): Ошибка: " & @ScriptDir & " (HKEY_LOCAL_MACHINE\XXXXXX\XXXXXXXXXXXXXXXXX\XXXXXXX\XXXXXXXXXXXXXXXXX\XXXXXXXX\XXXXXXX)")
	EndIf

	; Создаёт параметр с целым числом
	$result = RegWrite("HKEY_LOCAL_MACHINE\XXXXXX\XXXXXXXXXXXXXXXXX\XXXXXXX\XXXXXXXXXXXXXXXXX\XXXXXXXX\XXXXXXX\XXXXXXXXXXX", "XXXXXXX", "REG_DWORD", 0)
	If $result = 0 Then
		_FileWriteLog($hLogFile, "MainLoop(): Ошибка: " & @ScriptDir & " (HKEY_LOCAL_MACHINE\XXXXXX\XXXXXXXXXXXXXXXXX\XXXXXXX\XXXXXXXXXXXXXXXXX\XXXXXXXX\XXXXXXX\XXXXXXXXXXX, XXXXXXX:0)")
	EndIf
EndFunc

;Main
MainInit()
MainLoop()
MainExit()
;End Main
