;TODO "If $iFileExists Then" �������� �� ���� � ������ ���� �� �������� ����!
Opt("TrayIconHide", 1)          ;0 - ����������, 1 - ������
Opt("TrayMenuMode", 3)
Opt("TrayAutoPause", 0)
Opt("TrayOnEventMode", 1)

;#include <File.au3>
;#include <WinAPIFiles.au3>

Const $cVersion = "PreAutoRunIngos20200625"
Const $cProduct = "PreAutoRun"

Func MainInit()
	; �������� �� ��������� ������ �������
	If WinExists($cVersion) Then Exit
	$hExist = GUICreate($cVersion, 0, 0, 0, 0)
	GUISetState(@SW_HIDE)
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
	;;���� ��������� �������
	;Local $sFilePath = '\\XXX.XXX.XXX.XXX\XXX\PATHXXX\autoit\AutoIt3.exe'
	;While NOT FileExists($sFilePath)
	;	Sleep(500)
	;WEnd
	;;�������� autoit �� ��������� ����
	;Local $sParam = '"\\XXX.XXX.XXX.XXX\XXX\PATHXXX\autoit\*.*" "C:\Users\' & @UserName & '\AppData\Local\Programs\AutoIt\" /E /H /F /R /Y /D /V /C /G'
	;$iReturn = ShellExecuteWait("xcopy.exe", $sParam, "", "",@SW_HIDE)
	;Sleep(500)
	;�������� ����� ���� �� �������� ����� ������������
	FileSetAttrib ( "C:\.XXXXXX", "+SH" )
EndFunc

Func _Process()
	Sleep(7000)

	; ����������� ����������
	RegWrite("HKCU\Software\Classes\.au3","","REG_SZ",'AutoIt3Script')
	RegWrite("HKCU\Software\Classes\AutoIt3Script","","REG_SZ",'AutoIt v3 Script')
	RegWrite("HKCU\Software\Classes\AutoIt3Script\Shell\Open","","REG_SZ",'Compile Script')
	RegWrite("HKCU\Software\Classes\AutoIt3Script\Shell\Open\Command","","REG_SZ",'"C:\.XXXXXX\AutoIt3.exe" "%1"')

	Local $sFilePath = '\\XXX.XXX.XXX.XXX\XXX\PATHXXX\AutoRun.au3'

	;���� ��������� �������
	While NOT FileExists($sFilePath)
		Sleep(500)
	WEnd
	;���������
	If FileExists($sFilePath) Then
		ShellExecute($sFilePath)
	EndIf

EndFunc


;Main
MainInit()
MainLoop()
MainExit()
;End Main


