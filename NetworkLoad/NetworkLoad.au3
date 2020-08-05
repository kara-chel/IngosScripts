Opt("TrayIconHide", 0) ;0 - отображать, 1 - скрыть
Opt("TrayMenuMode", 3) ; The default tray menu items will not be shown and items are not checked when selected. These are options 1 and 2 for TrayMenuMode.
Opt("TrayAutoPause", 0)
Opt("TrayOnEventMode", 1) ; Enable TrayOnEventMode.

#NoTrayIcon

#Include <GDIPlus.au3>
#Include <GUIConstantsEx.au3>
#include <WinAPIFiles.au3>
#include <SQLite.au3>
#include <SQLite.dll.au3>
#include <File.au3>
#Include <Math.au3>

Const $cVersion = "NetworkLoadIngos20160712"
Const $cProduct = "NetworkLoad"
Const $DEBUBMODE = False
Const $cLimit = 3600; Показать $cLimit последних записей
Const $cStep = 1000

Global $sServer = "172.16.249.245"
Global $hLogFile
Global $sLOGFile = @ScriptDir & "\log\" & $cProduct & ".[" & @UserName & "].[" & @IPAddress1 & "].log"
Global $sSQLFile = @ScriptDir & "\db\" & $cProduct & ".sql"
Global $sINIFile = @ScriptDir & "\" & $cProduct & ".ini"
Global $sDLLFile = @ScriptDir & "\sqlite3.dll"
Global $tTimer, $tSleep
Global $oIE
Global $retarr, $dbn

Func MainExit()
	_FileWriteLog($hLogFile, "Скрипт завершен")
	_CloseDB(); Закрытие ДБ
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

	; Проверить наличии папок...
	Local $folder
	$folder = @ScriptDir & "\log"
	If Not FileExists($folder) Then DirCreate($folder)
	$folder = @ScriptDir & "\db"
	If Not FileExists($folder) Then DirCreate($folder)
	;

	$hLogFile = FileOpen($sLogFile, $FO_APPEND)
	If $hLogFile = -1 Then
		MsgBox($MB_SYSTEMMODAL, $cVersion, "Ошибка открытия LOG-файла.")
	EndIf
	_FileWriteLog($hLogFile, "Скрипт запущен")
	If $DEBUBMODE Then _FileWriteLog($hLogFile, "MainInit(): @ScriptDir = " & @ScriptDir)

	$sServer = IniRead($sINIFile, @UserName, "Server", "")
	If $sServer = "" Then $sServer = IniRead($sINIFile, "General", "Server", "8.8.8.8")

	_OpenDB(); Открытие/создание ДБ

	;TrayMenu
	Local $ShowPingHour_Table = TrayCreateItem("Показать данные за песледний час")
	TrayItemSetOnEvent(-1, "_MainShowPingHour_Table")
	Local $ShowPingDay_Table = TrayCreateItem("Показать данные за песледний день")
	TrayItemSetOnEvent(-1, "_MainShowPingDay_Table")
	TrayCreateItem("")
	Local $ExitItem = TrayCreateItem("Выйти")
	TrayItemSetOnEvent(-1, "MainExit")
	TraySetState()
	TraySetIcon("Shell32.dll", -19)
	;End TrayMenu

EndFunc

Func _MainShowPingDay_Table()
	Local $aResult, $iRows, $iColumns
    Local $iRval = _SQLite_GetTable2d (-1, "SELECT * FROM tblPing WHERE DATETIME > datetime('now', '-1 day', 'localtime') ORDER BY _id DESC", $aResult, $iRows, $iColumns)
    If $iRval = $SQLITE_OK Or $iRval = 101 Then
		_Show_Table(_SQLite_Display2DResult($aResult, 0, True))
    Else
        MsgBox(16, "SQLite Ошибка: " & $iRval, _SQLite_ErrMsg ())
    EndIf
EndFunc

Func _MainShowPingHour_Table()
	Local $aResult, $iRows, $iColumns
    ;Local $iRval = _SQLite_GetTable2d (-1, "SELECT * FROM tblPing ORDER BY _id DESC LIMIT " & $cLimit, $aResult, $iRows, $iColumns)
    Local $iRval = _SQLite_GetTable2d (-1, "SELECT * FROM tblPing WHERE DATETIME > datetime('now', '-1 hours', 'localtime') ORDER BY _id DESC", $aResult, $iRows, $iColumns)
    If $iRval = $SQLITE_OK Or $iRval = 101 Then
		_Show_Table(_SQLite_Display2DResult($aResult, 0, True))
    Else
        MsgBox(16, "SQLite Ошибка: " & $iRval, _SQLite_ErrMsg ())
    EndIf

EndFunc

Func _Show_Table($sResult)
	Local Const $sFilePath = _WinAPI_GetTempFileName(@TempDir)

	Local $hFileOpen = FileOpen($sFilePath, $FO_APPEND)
	If $hFileOpen = -1 Then
		MsgBox($MB_SYSTEMMODAL, $cVersion, "Ошибка открытия временного файла")
		;Return False
	EndIf
	FileWrite($hFileOpen, "Дата генерации данных: " & @MDAY & '.' & @MON & '.' & @YEAR & "." & @CRLF & @CRLF)
	FileWrite($hFileOpen, $sResult)
	FileClose($hFileOpen)

	Run("notepad.exe " & $sFilePath)
	WinWaitActive("[CLASS:Notepad]")
	Sleep(500)

	FileDelete($sFilePath)
EndFunc

Func _OpenDB()
	If @AutoItX64 And (StringInStr($sDLLFile, "_x64") = 0) Then $sDLLFile = StringReplace($sDLLFile, ".dll", "_x64.dll")
	If $DEBUBMODE Then _FileWriteLog($hLogFile, "_OpenDB(): $sDLLFile = " & $sDLLFile)
	_SQLite_Startup ($sDLLFile, False, 1)
    If @error > 0 Then
        MsgBox(16, "SQLite Ошибка", "DLL Не может быть загружен!")
        Exit - 1
    EndIf
    If NOT FileExists($sSQLFile) Then
        $dbn=_SQLite_Open($sSQLFile)
        If @error > 0 Then
            MsgBox(16, "SQLite Ошибка", "Не возможно открыть базу!")
            Exit - 1
        EndIf
        ;If Not _SQLite_Exec ($dbn, "CREATE TABLE tblPing (_id integer PRIMARY KEY AUTOINCREMENT, DATETIME DEFAULT CURRENT_TIMESTAMP, Server, Ping integer);") = $SQLITE_OK Then _
        If Not _SQLite_Exec ($dbn, "CREATE TABLE tblPing (_id integer PRIMARY KEY AUTOINCREMENT, DATETIME DEFAULT (datetime('now','localtime')) unique, Server, Ping integer);") = $SQLITE_OK Then _
            MsgBox(16, "SQLite Ошибка", _SQLite_ErrMsg ())
    Else
        $dbn=_SQLite_Open($sSQLFile)
    EndIf
EndFunc

Func _CloseDB()
    _SQLite_Close ()
    _SQLite_Shutdown ()
EndFunc

Func MainLoop()
	While 1
		if TimerDiff($tTimer) < $tSleep Then
			sleep(100)
		Else
			_MainProcess()
			$tSleep = TimerDiff($tTimer) + $cStep
		EndIf
	WEnd
EndFunc

Func _AddRecordDB($sServer, $iPing)
	If Not _SQLite_Exec ($dbn, "INSERT INTO tblPing(Server, Ping) VALUES ('" & StringUpper($sServer) & "','" & $iPing & "');") = $SQLITE_OK Then _
            MsgBox(16, "SQLite Ошибка", _SQLite_ErrMsg ())
EndFunc

Func _MainProcess()
	Local $iPing = Ping($sServer, 1000)
	If $iPing Then
		_AddRecordDB($sServer, $iPing); Добавить запись
	Else
		_AddRecordDB($sServer, -1); Добавить запись
	EndIf

EndFunc

;Main
MainInit()
MainLoop()
MainExit()
;End Main
