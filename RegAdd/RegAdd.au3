Opt("TrayIconHide", 0) ;0 - ����������, 1 - ������
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
	_FileWriteLog($hLogFile, "������ ��������")
	FileClose($hLogFile); ��������� LOG-����
	Exit
EndFunc

Func MainInit()
	; �������� �� ��������� ������ �������
	If WinExists($cVersion) Then Exit
	$hExist = GUICreate($cVersion, 0, 0, 0, 0)
	GUISetState(@SW_HIDE)
	; /

	; ������
	$tTimer = TimerInit()
	$tSleep = 0
	; /

	$hLogFile = FileOpen($sLogFile, $FO_APPEND)
	If $hLogFile = -1 Then
		MsgBox($MB_SYSTEMMODAL, $cVersion, "������ �������� LOG-�����")
	EndIf
	_FileWriteLog($hLogFile, "������ �������")
	TrayTip($cProduct, "������ �������", 3, 1)

	If $DEBUBMODE Then _FileWriteLog($hLogFile, "MainInit(): @ScriptDir = " & @ScriptDir)

	$aExceptionUserName = StringSplit(StringUpper(StringStripWS(IniRead($sINIFile, "Exception", "UserName", ""), 3)), ',')
	If _ArraySearch($aExceptionUserName, StringUpper(@UserName)) >= 0 Then
		_FileWriteLog($hLogFile, "MainInit(): Expection = " & StringUpper(@UserName))
		MainExit()
	EndIf

	;TrayMenu
	Local $exititem = TrayCreateItem("�����")
	TrayItemSetOnEvent(-1, "MainExit")
	TraySetState()
	TraySetIcon("Shell32.dll", -21)
	;End TrayMenu

EndFunc

Func MainLoop()
	_FileWriteLog($hLogFile, "MainLoop()")
	; ������ ������
	$result = RegWrite("HKEY_LOCAL_MACHINE\XXXXXX\XXXXXXXXXXXXXXXXX\XXXXXXX\XXXXXXXXXXXXXXXXX\XXXXXXXX\XXXXXXX")
	If $result = 0 Then
		_FileWriteLog($hLogFile, "MainLoop(): ������: " & @ScriptDir & " (HKEY_LOCAL_MACHINE\XXXXXX\XXXXXXXXXXXXXXXXX\XXXXXXX\XXXXXXXXXXXXXXXXX\XXXXXXXX\XXXXXXX)")
	EndIf

	; ������ �������� � ����� ������
	$result = RegWrite("HKEY_LOCAL_MACHINE\XXXXXX\XXXXXXXXXXXXXXXXX\XXXXXXX\XXXXXXXXXXXXXXXXX\XXXXXXXX\XXXXXXX\XXXXXXXXXXX", "XXXXXXX", "REG_DWORD", 0)
	If $result = 0 Then
		_FileWriteLog($hLogFile, "MainLoop(): ������: " & @ScriptDir & " (HKEY_LOCAL_MACHINE\XXXXXX\XXXXXXXXXXXXXXXXX\XXXXXXX\XXXXXXXXXXXXXXXXX\XXXXXXXX\XXXXXXX\XXXXXXXXXXX, XXXXXXX:0)")
	EndIf
EndFunc

;Main
MainInit()
MainLoop()
MainExit()
;End Main
