Opt("TrayIconHide", 0) ;0 - ����������, 1 - ������
Opt("TrayMenuMode", 3) ; The default tray menu items will not be shown and items are not checked when selected. These are options 1 and 2 for TrayMenuMode.
Opt("TrayAutoPause", 0)
Opt("TrayOnEventMode", 1) ; Enable TrayOnEventMode.

#NoTrayIcon

;#Include <GDIPlus.au3>
;#Include <GUIConstantsEx.au3>
;#include <WinAPIFiles.au3>
#include <File.au3>
;#Include <Math.au3>

Const $cVersion = "MenuIngos20200710"
Const $cProduct = "MenuIngos"
Const $DEBUBMODE = False
Const $cStep = 100

Global $hLogFile
Global $sLOGFile = @ScriptDir & "\log\" & $cProduct & ".[" & @UserName & "].[" & @IPAddress1 & "].log"
Global $sINIFile = @ScriptDir & "\" & $cProduct & ".ini"
Global $tTimer, $tSleep
Global $aAdmin


Func MainExit()
	; ������ ��������� ������������ ����� �����.
	Local $iFind = _ArraySearch($aAdmin, StringUpper(@UserName))
	if  $iFind >= 0 Then
		FileClose($hLogFile)
		Exit
	Else
		MsgBox($MB_SYSTEMMODAL, "�����", "�� �� ������ ���� �� �������� ���� ���������.")
		_FileWriteLog($hLogFile, "MainExit(): ������� ������� ���������.")

	EndIf
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

	$hLogFile = FileOpen($sLOGFile, $FO_APPEND)
	If $hLogFile = -1 Then
		MsgBox($MB_SYSTEMMODAL, $cVersion, "������ �������� LOG-�����.")
	EndIf
	_FileWriteLog($hLogFile, "������ �������")
	If $DEBUBMODE Then _FileWriteLog($hLogFile, "MainInit(): @ScriptDir = " & @ScriptDir)


	$aAdmin = StringSplit(StringUpper(StringStripWS(IniRead($sINIFile, "General", "Admin", ''), 8)), ',')

	;TrayMenu
	TrayCreateItem("������������� ���")
	TrayItemSetOnEvent(-1, "_Menu0")

	TrayCreateItem("������������� ����")
	TrayItemSetOnEvent(-1, "_Menu1")

	TrayCreateItem("")

	TrayCreateItem("�������� ����������/��������")
	TrayItemSetOnEvent(-1, "_Menu2")

	TrayCreateItem("")

	TrayCreateItem("��������� �����")
	TrayItemSetOnEvent(-1, "_Menu3")

	TrayCreateItem("������������ ����������")
	TrayItemSetOnEvent(-1, "_Menu3")

	TrayCreateItem("")
	Local $ExitItem = TrayCreateItem("�����")
	TrayItemSetOnEvent(-1, "MainExit")
	TraySetState()
	TraySetIcon("Shell32.dll", -36)
	;End TrayMenu

EndFunc


Func _Menu0()
	_FileWriteLog($hLogFile, "Click _Menu0")
	Local $YesNo = MsgBox(1, "������������� ���", "�� ������� ��� ������ �������������� ���?")
	_FileWriteLog($hLogFile, "Click _Menu0: " & $YesNo)
	if $YesNo = 1 Then
		$iReturn = ShellExecuteWait(@ScriptDir & "\cmd\������������� ���.cmd", "", "", "",@SW_SHOW)
	EndIf
EndFunc

Func _Menu1()
	_FileWriteLog($hLogFile, "Click _Menu1")
	Local $YesNo = MsgBox(1, "������������� ����", "�� ������� ��� ������ �������������� ����? ����� ������������� ����� ����������� ���������� ����. ����� ���������� ��������� �� ����.")
	_FileWriteLog($hLogFile, "Click _Menu1: " & $YesNo)
	if $YesNo = 1 Then
		$iReturn = ShellExecuteWait(@ScriptDir & "\cmd\������������� ����.cmd", "", "", "",@SW_SHOW)
	EndIf
EndFunc

Func _Menu2()
	_FileWriteLog($hLogFile, "Click _Menu2")
	$iReturn = ShellExecuteWait(@ScriptDir & "\cmd\�������� ��������.cmd", "", "", "",@SW_SHOW)
EndFunc

Func _Menu3()
	_FileWriteLog($hLogFile, "Click _Menu3")
	$iReturn = ShellExecuteWait("C:\Windows\System32\taskmgr.exe", "", "", "",@SW_SHOW)
EndFunc

Func _Menu4()
	_FileWriteLog($hLogFile, "Click _Menu4")
	Local $YesNo = MsgBox(1, "�������������", "�� ������� ��� ������ ������������� ���������?")
	_FileWriteLog($hLogFile, "Click _Menu4: " & $YesNo)
	if $YesNo = 1 Then
		$iReturn = ShellExecuteWait("C:\Windows\System32\shutdown.exe", "/r /f", "", "",@SW_HIDE)
	EndIf

EndFunc

Func MainLoop()
	While 1
		sleep(100)
	WEnd
EndFunc

Func _MainProcess()
EndFunc

;Main
MainInit()
MainLoop()
MainExit()
;End Main
