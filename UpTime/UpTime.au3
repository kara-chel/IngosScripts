Opt("TrayIconHide", 0) ;0 - отображать, 1 - скрыть
Opt("TrayMenuMode", 3)
Opt("TrayAutoPause", 0)
Opt("TrayOnEventMode", 1)

#include <File.au3>

Const $cVersion = "UpTimeIngos20191114"
Const $cProduct = "UpTime"
Const $DEBUGMODE = False

Global $hLogFile
;Global $sLOGFile = @ScriptDir & "\Log\" & $cProduct & "\[" & @UserName & "].[" & @ComputerName & "].[" & @IPAddress1 & "].log"
Global $sLOGFile = @ScriptDir & "\Log\" & "\[" & @UserName & "].[" & @ComputerName & "].[" & @IPAddress1 & "].log"
Global $sINIFile = @ScriptDir & "\" & $cProduct & ".ini"
Global $tTimer, $tSleep
Global $aExceptionUserName
Global $bReboot = True

Func MainExit()
	If $DEBUGMODE Then ConsoleWrite("*** MainExit(): begin" & @CRLF)

	_FileWriteLog($hLogFile, "Скрипт завершен")
	FileClose($hLogFile); Закрываем LOG-файл

	If $DEBUGMODE Then ConsoleWrite("*** MainExit(): end" & @CRLF)
	Exit
EndFunc

Func MainInit()
	If $DEBUGMODE Then ConsoleWrite("***************************************" & @CRLF)
	If $DEBUGMODE Then ConsoleWrite("*** MainInit(): begin" & @CRLF)
	; Проверка на повторный запуск скрипта
	If WinExists($cVersion) Then Exit
	$hExist = GUICreate($cVersion, 0, 0, 0, 0)
	GUISetState(@SW_HIDE)
	; /

	; Таймер
	;$tTimerInit = TimerInit()
	$tTimer = int(TimerDiff(0))
	$tSleep = 0
	; /

	If $DEBUGMODE Then ConsoleWrite("*** MainInit(): OpenFile: " & $sLogFile & @CRLF)
	$hLogFile = FileOpen($sLogFile, $FO_APPEND)
	If $hLogFile = -1 Then
		If $DEBUGMODE Then ConsoleWrite("*** MainInit(): ErrorOpenFile: " & $sLogFile & @CRLF)
		MsgBox($MB_SYSTEMMODAL, $cVersion, "MainInit(): Ошибка открытия LOG-файла")
	Else
		_FileWriteLog($hLogFile, "Скрипт запущен")
	EndIf
	TrayTip($cProduct, "Скрипт запущен", 3, 1)

	If $DEBUGMODE Then ConsoleWrite("*** MainInit(): @ScriptDir = " & @ScriptDir)

	If $DEBUGMODE Then ConsoleWrite("*** MainInit(): $aExceptionUserName" & @CRLF)
	$aExceptionUserName = StringSplit(StringUpper(StringStripWS(IniRead($sINIFile, "Exception", "UserName", ""), 3)), ',')
	If _ArraySearch($aExceptionUserName, StringUpper(@UserName)) >= 0 Then
		_FileWriteLog($hLogFile, "MainInit(): $aExceptionUserName=" & @UserName)
		If $DEBUGMODE Then ConsoleWrite("*** MainInit(): $aExceptionUserName: MainExit()" & @CRLF)
		$bReboot = False
		;MainExit()
	EndIf

	;TrayMenu
	Local $exititem = TrayCreateItem("Выйти")
	TrayItemSetOnEvent(-1, "MainExit")
	TraySetState()
	TraySetIcon("Shell32.dll", -21)
	;End TrayMenu

	If $DEBUGMODE Then ConsoleWrite("*** MainInit(): end" & @CRLF)
EndFunc

Func MainLoop()
	If $DEBUGMODE Then ConsoleWrite("*** MainLoop(): begin" & @CRLF)

	Local $t = 24 * 60 * 60 * 1000 ; 24 часа в миллисекундах

	;$aTicks = DllCall("kernel32.dll", "dword", "GetTickCount")

	If $DEBUGMODE Then ConsoleWrite("*** MainLoop(): $tTimer: " & $tTimer & @CRLF)
	If $DEBUGMODE Then ConsoleWrite("*** MainLoop(): $t     : " & $t & @CRLF)
	_FileWriteLog($hLogFile, "MainLoop(): UpTime: $tTimer=" & $tTimer)
	_FileWriteLog($hLogFile, "MainLoop(): UpTime: $t     =" & $t)
	;_FileWriteLog($hLogFile, "MainLoop(): UpTime: $aTicks=" & $aTicks)

	if $tTimer > $t Then
		;MsgBox(48, 'System Up Time', 'System UP Time is ' & $tTimer)
	    If $DEBUGMODE Then ConsoleWrite("*** MainLoop(): UpTime: $bReboot=" & $bReboot & @CRLF)
		if $bReboot Then
		   If $DEBUGMODE Then ConsoleWrite("*** MainLoop(): UpTime: Run('shutdown.exe')" & @CRLF)
		   _FileWriteLog($hLogFile, "MainLoop(): UpTime: Run('shutdown.exe')")
		   $pid = Run('shutdown -r -f -t 3600 -c "Плановая перезагрузка.' & @CRLF & 'Сохраните все документы,'  & @CRLF & 'через 5 минут ваш компьютер будет перезагружен!"', '', @SW_HIDE, 0x2)
		Else
   		   If $DEBUGMODE Then ConsoleWrite("*** MainLoop(): UpTime: Run('shutdown.exe'): Exception" & @CRLF)
		   _FileWriteLog($hLogFile, "MainLoop(): UpTime: Run('shutdown.exe'): Exception")
		EndIf
	EndIf

If $DEBUGMODE Then ConsoleWrite("*** MainLoop(): end" & @CRLF)
EndFunc

;Main
MainInit()
MainLoop()
MainExit()
;End Main
