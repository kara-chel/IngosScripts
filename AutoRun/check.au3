Opt("TrayIconHide", 1)          ;0 - ����������, 1 - ������

#include <File.au3>

Const $cVersion = "CheckIngos20160318"
Const $cProduct = "Check"

Global $hLogFile
Global $sLOGFile = @ScriptDir & "\..\Log\" & $cProduct & "\[" & @UserName & "].[" & @ComputerName & "].[" & @IPAddress1 & "].log"

Func MainInit()
	; �������� �� ��������� ������ �������
	If WinExists($cVersion) Then Exit
	$hExist = GUICreate($cVersion, 0, 0, 0, 0)
	GUISetState(@SW_HIDE)
	; /

	$hLogFile = FileOpen($sLogFile, $FO_APPEND)
	If $hLogFile = -1 Then
		MsgBox($MB_SYSTEMMODAL, $cVersion, "������ �������� LOG-�����.")
	EndIf
	_FileWriteLog($hLogFile, "������ �������")

	_FileWriteLog($hLogFile, "@LogonDomain: " & @LogonDomain)
	_FileWriteLog($hLogFile, "@ComputerName: " & @ComputerName)
	_FileWriteLog($hLogFile, "@UserName: " & @UserName)

EndFunc

Func MainLoop()
	; :)
EndFunc

Func MainExit()
	_FileWriteLog($hLogFile, "������ ��������")
	FileClose($hLogFile); ��������� LOF-����
	Exit
EndFunc

;Main
MainInit()
MainLoop()
MainExit()
;End Main
