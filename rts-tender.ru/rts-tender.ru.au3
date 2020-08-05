Opt("TrayIconHide", 0) ;0 - отображать, 1 - скрыть
Opt("TrayMenuMode", 3) ; The default tray menu items will not be shown and items are not checked when selected. These are options 1 and 2 for TrayMenuMode.
Opt("TrayAutoPause", 0)
Opt("TrayOnEventMode", 1) ; Enable TrayOnEventMode.

#NoTrayIcon

#include <WinAPIFiles.au3>
#include <SQLite.au3>
#include <SQLite.dll.au3>
#include <File.au3>

Const $cVersion = "RtsTenderRuIngos20190918"
Const $cProduct = "rts-tender.ru"

Global $iMin
Global $iMax
Global $iNext

Global $hLogFile
Global $sLOGFile = @ScriptDir & "\log\" & $cProduct & ".[" & @UserName & "].[" & @IPAddress1 & "].log"
Global $sSQLFile = @ScriptDir & "\db\rts-tender.ru.[" & @UserName & "].sql"
Global $sINIFile = @ScriptDir & "\" & $cProduct & ".ini"
Global $sDLLFile = @ScriptDir & "\sqlite3.dll"
Global $tTimer, $tSleep = 0
Global $oIE
Global $retarr, $dbn
Global $DEBUGMODE = True
Global $iShowLimit; Показать $iShowLimit последних записей

Func MainExit()
	_FileWriteLog($hLogFile, "Скрипт завершен")
	_CloseDB(); Закрытие ДБ
	FileClose($hLogFile); Закрываем LOG-файл
	Exit
EndFunc

Func _MainShowTender()
	Local $aResult, $iRows, $iColumns
    Local $iRval = _SQLite_GetTable2d (-1, "SELECT * FROM tblTender WHERE UserName='" & StringUpper(@UserName) & "' ORDER BY _id DESC LIMIT " & $iShowLimit, $aResult, $iRows, $iColumns)
    If $iRval = $SQLITE_OK Then
		$sResult = _SQLite_Display2DResult($aResult, 0, True)

		Local Const $sFilePath = _WinAPI_GetTempFileName(@TempDir)

		Local $hFileOpen = FileOpen($sFilePath, $FO_APPEND)
		If $hFileOpen = -1 Then
			MsgBox($MB_SYSTEMMODAL, $cVersion, "Ошибка открытия временного файла")
			;Return False
		EndIf
		FileWrite($hFileOpen, "База найденых тендеров на rts-tender.ru от " & @MDAY & '.' & @MON & '.' & @YEAR & "." & @CRLF & @CRLF)
		FileWrite($hFileOpen, $sResult)
		FileClose($hFileOpen)

		Run("notepad.exe " & $sFilePath)
		WinWaitActive("[CLASS:Notepad]")
		Sleep(500)

		FileDelete($sFilePath)
    Else
        MsgBox(16, "SQLite Ошибка: " & $iRval, _SQLite_ErrMsg ())
    EndIf

EndFunc

Func _MainExec()
	_MainProcess()
	;$tSleep = TimerDiff($tTimer) + Random($iMax, $iMax) * 1000
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
		MsgBox($MB_SYSTEMMODAL, $cVersion, "Ошибка открытия LOG-файла.")
	EndIf
	_FileWriteLog($hLogFile, "Скрипт запущен")
	;TrayTip($cProduct, "Скрипт запущен", 3, 1)

	If $DEBUGMODE Then _FileWriteLog($hLogFile, "MainInit(): @ScriptDir = " & @ScriptDir)

	$iMin = IniRead($sINIFile, @UserName, "MIN", "")
	If $iMin = "" Then $iMin = IniRead($sINIFile, "General", "MIN", '30')
	$iMax = IniRead($sINIFile, @UserName, "MAX", "")
	If $iMax = "" Then $iMax = IniRead($sINIFile, "General", "MAX", '120')
	;*****
	$iShowLimit = IniRead($sINIFile, @UserName, "ShowLimit", "")
	If $iShowLimit = "" Then $iShowLimit = IniRead($sINIFile, "General", "ShowLimit", '100')
	;*****
	Local $sTmp = IniRead($sINIFile, @UserName, "DEBUGMODE", "")
	If $sTmp = "" Then $sTmp = IniRead($sINIFile, "General", "DEBUGMODE", "False")
	If $sTmp = "True" Or $sTmp = "TRUE" Then
		$DEBUGMODE = True
	Else
		$DEBUGMODE = False
	EndIf

	If $DEBUGMODE Then _FileWriteLog($hLogFile, "MainInit(): $DEBUGMODE = " & $DEBUGMODE)
	If $DEBUGMODE Then _FileWriteLog($hLogFile, "MainInit(): $iMin = " & $iMin)
	If $DEBUGMODE Then _FileWriteLog($hLogFile, "MainInit(): $iMax = " & $iMax)

	_OpenDB(); Открытие/создание ДБ

	;TrayMenu
	Local $ExecItem = TrayCreateItem("Принудительно запустить скрипт")
	TrayItemSetOnEvent(-1, "_MainExec")
	Local $ShowDBItem = TrayCreateItem("Показать тендеры")
	TrayItemSetOnEvent(-1, "_MainShowTender")
	TrayCreateItem("")
	Local $ExitItem = TrayCreateItem("Выйти")
	TrayItemSetOnEvent(-1, "MainExit")
	TraySetState()
	TraySetIcon("Shell32.dll", -14)
	;End TrayMenu

EndFunc

Func _OpenDB()
	If @AutoItX64 And (StringInStr($sDLLFile, "_x64") = 0) Then $sDLLFile = StringReplace($sDLLFile, ".dll", "_x64.dll")
	If $DEBUGMODE Then _FileWriteLog($hLogFile, "_OpenDB(): $sDLLFile = " & $sDLLFile)
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
        If Not _SQLite_Exec ($dbn, "CREATE TABLE tblTender (_id integer PRIMARY KEY AUTOINCREMENT, UserName, Title, URL, Date, Time);") = $SQLITE_OK Then _
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
			$tSleep = TimerDiff($tTimer) + $iNext * 1000
		EndIf
	WEnd
EndFunc

Func _MainProcess()
	;ShellExecute(@ScriptDir & "\" & $cProduct &".ext.au3")
	Local $iFileExists = FileExists(@ScriptDir & "\rts-tender.ru.ext.au3")
	If $iFileExists Then
		If $DEBUGMODE Then _FileWriteLog($hLogFile, "_MainProcess: Execute AU3")
		ShellExecute(@ScriptDir & "\rts-tender.ru.ext.au3")
	Else
		If $DEBUGMODE Then _FileWriteLog($hLogFile, "_MainProcess: Execute EXE")
		ShellExecute(@ScriptDir & "\rts-tender.ru.ext.exe")
	EndIf
	$iNext = Int(Random($iMin, $iMax))
	;TrayTip($cProduct, "Сайт проверен, следующая проверка через " & $iNext & " секунд", 3, 1)
EndFunc

;Main
MainInit()
MainLoop()
MainExit()
;End Main
