;TODO "If $iFileExists Then" заменить на цикл с паузой пока не появится файл!
Opt("TrayIconHide", 1)          ;0 - отображать, 1 - скрыть
Opt("TrayMenuMode", 3)
Opt("TrayAutoPause", 0)
Opt("TrayOnEventMode", 1)

;#include <File.au3>
;#include <WinAPIFiles.au3>

Const $cVersion = "PreAutoRunIngos20200625"
Const $cProduct = "PreAutoRun"

Func MainInit()
	; Проверка на повторный запуск скрипта
	If WinExists($cVersion) Then Exit
	AutoItWinSetTitle($cVersion)
	;$hExist = GUICreate($cVersion, 0, 0, 0, 0)
	;GUISetState(@SW_HIDE)
	; /

EndFunc

Func MainLoop()
	_Install()
	_Process()
EndFunc

Func MainExit()
	Exit
EndFunc

Func _Install()
	;;Ждем появления ресурса
	;Local $sFilePath = '\\172.16.1.2\doc\Scripts\autoit\AutoIt3.exe'
	;While NOT FileExists($sFilePath)
	;	Sleep(500)
	;WEnd
	;;Копируем autoit на локальный диск
	;Local $sParam = '"\\172.16.1.2\doc\Scripts\autoit\*.*" "C:\Users\' & @UserName & '\AppData\Local\Programs\AutoIt\" /E /H /F /R /Y /D /V /C /G'
	;$iReturn = ShellExecuteWait("xcopy.exe", $sParam, "", "",@SW_HIDE)
	;Sleep(500)
	FileSetAttrib ( "C:\.autoit", "+SH" )
EndFunc

Func _Process()
	Sleep(7000)

	; Прописываем ассоциации
	RegWrite("HKCU\Software\Classes\.au3","","REG_SZ",'AutoIt3Script')
	RegWrite("HKCU\Software\Classes\AutoIt3Script","","REG_SZ",'AutoIt v3 Script')
	RegWrite("HKCU\Software\Classes\AutoIt3Script\Shell\Open","","REG_SZ",'Compile Script')
	;RegWrite("HKCU\Software\Classes\AutoIt3Script\Shell\Open\Command","","REG_SZ",'"C:\Users\' & @UserName & '\AppData\Local\Programs\AutoIt\AutoIt3.exe" "%1"')
	RegWrite("HKCU\Software\Classes\AutoIt3Script\Shell\Open\Command","","REG_SZ",'"C:\.autoit\AutoIt3.exe" "%1"')

	Local $sFilePath = '\\172.16.1.2\Doc\Scripts\AutoRun.au3'

	;Ждем появления ресурса
	While NOT FileExists($sFilePath)
		Sleep(500)
	WEnd
	;Запускаем
	If FileExists($sFilePath) Then
		ShellExecute($sFilePath)
	EndIf

EndFunc


;Main
MainInit()
MainLoop()
MainExit()
;End Main


