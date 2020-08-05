Opt("TrayIconHide", 0) ;0 - ����������, 1 - ������
Opt("TrayMenuMode", 3)
Opt("TrayAutoPause", 0)
Opt("TrayOnEventMode", 1)

#include <File.au3>
#include <FileOperations.au3>

Const $cVersion = "FindMP3Ingos20160419"
Const $cProduct = "FindMP3"
Const $DEBUBMODE = True

Global $hLogFile
Global $sLOGFile = @ScriptDir & "\..\Log\" & $cProduct & "\[" & @UserName & "].[" & @ComputerName & "].[" & @IPAddress1 & "].log"
;Global $sINIFile = @ScriptDir & "\" & $cProduct & ".ini"
Global $aExceptionUserName

Func MainExit()
	_FileWriteLog($hLogFile, "������ ��������")
	FileClose($hLogFile); ��������� LOG-����
	Exit
EndFunc

Func MainInit()
	If $DEBUBMODE Then _FileWriteLog($hLogFile, "MainInit()")
	; �������� �� ��������� ������ �������
	If WinExists($cVersion) Then Exit
	$hExist = GUICreate($cVersion, 0, 0, 0, 0)
	GUISetState(@SW_HIDE)
	; /

	$hLogFile = FileOpen($sLogFile, $FO_APPEND)
	If $hLogFile = -1 Then
		MsgBox($MB_SYSTEMMODAL, $cVersion, "������ �������� LOG-�����")
	EndIf
	_FileWriteLog($hLogFile, "������ �������")
	TrayTip($cProduct, "������ �������", 3, 1)

	If $DEBUBMODE Then _FileWriteLog($hLogFile, "MainInit(): @ScriptDir = " & @ScriptDir)

	$aExceptionUserName = StringSplit(StringUpper(StringStripWS(IniRead($sINIFile, "Exception", "UserName", ""), 3)), ',')
	If _ArraySearch($aExceptionUserName, StringUpper(@UserName)) >= 0 Then MainExit()

	;TrayMenu
	Local $exititem = TrayCreateItem("�����")
	TrayItemSetOnEvent(-1, "MainExit")
	TraySetState()
	TraySetIcon("Shell32.dll", -23)
	;End TrayMenu

EndFunc

Func MainLoop()
	If $DEBUBMODE Then _FileWriteLog($hLogFile, "MainLoop()")
	_MainProcess()
EndFunc

Func _MainProcess()
	If $DEBUBMODE Then _FileWriteLog($hLogFile, "_MainProcess()")

	Local $FileList, $File

	$FileList = _FO_FileSearch("C:\", "*.mp3", True, 255)
	For $File In $FileList
		$File = StringStripWS($File, $STR_STRIPALL)
		 _FileWriteLog($hLogFile, "_MainProcess(): ������: " & $File)
	Next
	_FileWriteLog($hLogFile, "_MainProcess(): ����� ������ �������: " & $FileList[0])

	;Sleep(1500)
EndFunc


;Main
MainInit()
MainLoop()
MainExit()
;End Main
