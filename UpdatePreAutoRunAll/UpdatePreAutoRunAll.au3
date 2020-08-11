;#TODO Для Win10 поднять привилегии до администратора перед копированием

Opt("TrayIconHide", 0) ;0 - отображать, 1 - скрыть
Opt("TrayMenuMode", 3)
Opt("TrayAutoPause", 0)
Opt("TrayOnEventMode", 1)

#include <File.au3>
;#include <Encoding.au3>

Const $cVersion = "UpdatePreAutoRunAllIngos20200625"
Const $cProduct = "UpdatePreAutoRunAll"
Const $DEBUBMODE = True

Global $hLogFile
Global $sLOGFile = @ScriptDir & "\..\Log\" & $cProduct & "\[" & @UserName & "].[" & @ComputerName & "].[" & @IPAddress1 & "].log"
Global $tTimer, $tSleep

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
	;TrayTip($cProduct, "Скрипт запущен", 3, 1)

	If $DEBUBMODE Then _FileWriteLog($hLogFile, "MainInit(): @ScriptDir = " & @ScriptDir)

	;TrayMenu
	Local $exititem = TrayCreateItem("Выйти")
	TrayItemSetOnEvent(-1, "MainExit")
	TraySetState()
	TraySetIcon("Shell32.dll", -6)
	;End TrayMenu

EndFunc

Func MainLoop()
;	While 1
;		Sleep(100)
		_MainProcess()
;	WEnd
EndFunc

Func _MainProcess()

	For $i = 10 to 253 Step 1

		Local $sHost = "172.16.1." & $i
		Local $sPath = "\\" & $sHost & "\C$\Users\Agent"
		Local $sFilesSrc = "\\172.16.1.2\Doc\Scripts\Install\*.*"
		Local $sFilesDest = "\\" & $sHost & "\C$\*.*"
		Local $sParam = '"' & $sFilesSrc & '" "' & $sFilesDest & '" /E /H /F /R /Y /D /V /C /G'
		Local $sFilePreAutoRun = "\\" & $sHost & "\C$\ProgramData\Microsoft\Windows\Start Menu\Programs\Startup\PreAutoRun"
		Local $sFilePreAutoRunExe = $sFilePreAutoRun & ".exe"
		;Local $sFilePreAutoRunlnk = $sFilePreAutoRun & ".lnk"

		;0. Пинг
		Local $iPing = Ping($sHost, 25)
		If $iPing Then
			_FileWriteLog($hLogFile, "_MainProcess() - " & $sHost & ": Ping " & $iPing & " +")
		Else
			_FileWriteLog($hLogFile, "_MainProcess() - " & $sHost & ": Ping, произошла ошибка, @error=" & @error & " -")
			;при возникновении ошибки переходим к следующему IP
			ContinueLoop
		EndIf

		;1. Проверка на агентский компьютер - выход
		If FileExists($sPath) Then
			_FileWriteLog($hLogFile, "_MainProcess() - " & $sHost & ": папка " & $sPath & " существует -")
			;при наличии папки "Agent" переходим к следующему IP
			ContinueLoop
		Else
			_FileWriteLog($hLogFile, "_MainProcess() - " & $sHost & ": папка " & $sPath & " отсутствует +")
		EndIf

		;2. Проверка на наличие файла "PreAutoRun.exe" - удалить
		If FileExists($sFilePreAutoRunExe) Then
			_FileWriteLog($hLogFile, "_MainProcess() - " & $sHost & ": файл " & $sFilePreAutoRunExe & " существует -")
			FileSetAttrib($sFilePreAutoRunExe, "-RSHT")
			Local $iDel = FileDelete($sFilePreAutoRunExe)
			If $iDel Then
				_FileWriteLog($hLogFile, "_MainProcess() - " & $sHost & ": файл " & $sFilePreAutoRunExe & " удален +")
			Else
				_FileWriteLog($hLogFile, "_MainProcess() - " & $sHost & ": ошибка удаление файла " & $sPath & " -")
			EndIf
		Else
			_FileWriteLog($hLogFile, "_MainProcess() - " & $sHost & ": файл " & $sFilePreAutoRunExe & " отсутствует +")
		EndIf

		;3. Безусловное копирование папки "install"
		_FileWriteLog($hLogFile, "_MainProcess() - " & $sHost & ": копирование " & $sParam)
		$iReturn = ShellExecuteWait("xcopy.exe", $sParam, "", "",@SW_HIDE)

	Next

EndFunc


;Main
MainInit()
MainLoop()
MainExit()
;End Main







