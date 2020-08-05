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
#include <SendMessage.au3>

Const $cVersion = "NetworkLoadClientIngos20160712"
Const $cProduct = "NetworkLoadClient"
Const $DEBUBMODE = False
Const $cLimit = 3600; Показать $cLimit последних записей
Const $cStep = 1000

Global $hLogFile
Global $sLOGFile = @ScriptDir & "\log\" & $cProduct & ".[" & @UserName & "].[" & @IPAddress1 & "].log"
Global $sSQLFile = @ScriptDir & "\db\" & "NetworkLoad.sql"
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
	Local $ShowDBPing1H_GUI = TrayCreateItem("Показать график за последний час")
	TrayItemSetOnEvent(-1, "_MainShowPing1H_GUI")
	Local $ShowDBPing30M_GUI = TrayCreateItem("Показать график за последниe 30 минут")
	TrayItemSetOnEvent(-1, "_MainShowPing30M_GUI")
	Local $ShowDBPing20M_GUI = TrayCreateItem("Показать график за последниt 20 минут")
	TrayItemSetOnEvent(-1, "_MainShowPing20M_GUI")
	Local $ShowDBPing10M_GUI = TrayCreateItem("Показать график за последниt 10 минут")
	TrayItemSetOnEvent(-1, "_MainShowPing10M_GUI")
	Local $ShowPingHour_Table = TrayCreateItem("Показать данные за песледний час")
	TrayItemSetOnEvent(-1, "_MainShowPingHour_Table")
	Local $ShowPingDay_Table = TrayCreateItem("Показать данные за песледний день")
	TrayItemSetOnEvent(-1, "_MainShowPingDay_Table")
	TrayCreateItem("")
	Local $ExitItem = TrayCreateItem("Выйти")
	TrayItemSetOnEvent(-1, "MainExit")
	TraySetState()
	TraySetIcon("Shell32.dll", -18) ;18 19
	;End TrayMenu

EndFunc

Dim $aPoints[$cLimit][2]
Global $iI
Func _cb($aRow)

	if $iI < $cLimit And $iI >= 0 Then
		$aPoints[$iI][0] = $iI
		$aPoints[$iI][1] = $aRow[3]
		ConsoleWrite($aRow[3] & @LF)
	EndIf
	$iI = $iI + 1
EndFunc   ;==>_cb

Func _MainShowPing1H_GUI()
	$iI = -1
	Local $d = _SQLite_Exec(-1, "SELECT * FROM tblPing WHERE DATETIME > datetime('now', '-1 hours', 'localtime') ORDER BY _id ASC LIMIT " & $cLimit, "_cb") ; _cb будет вызвана для каждой строки
	_ViewGraph($aPoints)
EndFunc

Func _MainShowPing30M_GUI()
	$iI = -1
    Local $d = _SQLite_Exec(-1, "SELECT * FROM tblPing WHERE DATETIME > datetime('now', '-30 minutes', 'localtime') ORDER BY _id ASC LIMIT " & $cLimit, "_cb") ; _cb будет вызвана для каждой строки
	_ViewGraph($aPoints)
EndFunc

Func _MainShowPing20M_GUI()
	$iI = -1
    Local $d = _SQLite_Exec(-1, "SELECT * FROM tblPing WHERE DATETIME > datetime('now', '-20 minutes', 'localtime') ORDER BY _id ASC LIMIT " & $cLimit, "_cb") ; _cb будет вызвана для каждой строки
	_ViewGraph($aPoints)
EndFunc

Func _MainShowPing10M_GUI()
	$iI = -1
    Local $d = _SQLite_Exec(-1, "SELECT * FROM tblPing WHERE DATETIME > datetime('now', '-10 minutes', 'localtime') ORDER BY _id ASC LIMIT " & $cLimit, "_cb") ; _cb будет вызвана для каждой строки
	_ViewGraph($aPoints)
EndFunc

Func _ViewGraph($aPoints)
	Const $WX = 1000
	Const $WY = 300

    Local $hForm, $Pic, $hPic
    Local $XScale, $YScale, $Xi, $Yi, $Xp, $Yp, $XOffset, $YOffset, $Xmin = $aPoints[0][0], $Ymin = $aPoints[0][1], $Xmax = $Xmin, $Ymax = $Ymin
    Local $hBitmap, $hObj, $hGraphic, $hImage, $hBrush, $hPen

	For $i = 1 To UBound($aPoints) - 1
        If $aPoints[$i][0] < $Xmin Then
            $Xmin = $aPoints[$i][0]
        Else
            If $aPoints[$i][0] > $Xmax Then
                $Xmax = $aPoints[$i][0]
            EndIf
        EndIf
        If $aPoints[$i][1] < $Ymin Then
            $Ymin = $aPoints[$i][1]
        Else
            If $aPoints[$i][1] > $Ymax Then
                $Ymax = $aPoints[$i][1]
            EndIf
        EndIf
    Next

    $XScale = $WX / $Xmax - $Xmin
    $YScale = 1
    $XOffset = 0
    $YOffset = $WY

    _GDIPlus_Startup()
    $hBitmap = _WinAPI_CreateBitmap($WX+1, $WY+1, 1, 32)
    $hImage = _GDIPlus_BitmapCreateFromHBITMAP($hBitmap)
    _WinAPI_DeleteObject($hBitmap)
    $hGraphic = _GDIPlus_ImageGetGraphicsContext($hImage)
    $hBrush = _GDIPlus_BrushCreateSolid(0xFFFFFFFF)
    _GDIPlus_GraphicsFillRect($hGraphic, 0, 0, $WX+1, $WY+1, $hBrush)
    _GDIPlus_GraphicsSetSmoothingMode($hGraphic, 2)

	; Рисуем координатнае плоскости
	$hPen = _GDIPlus_PenCreate(0xFF0000FF)
	; Рисуем координатную ось Y
    _GDIPlus_GraphicsDrawLine($hGraphic, $XOffset, 0, $XOffset, $WY + 1, $hPen)
	; Рисуем координатную ось X
    _GDIPlus_GraphicsDrawLine($hGraphic, 0, $YOffset, $WX + 1, $YOffset, $hPen)
    _GDIPlus_PenDispose($hPen)

	; рисуем линию относительного лимита
	$hPen = _GDIPlus_PenCreate(0xFF00FF00)
    _GDIPlus_GraphicsDrawLine($hGraphic, 0, $YOffset-100, $WX + 1, $YOffset-100, $hPen)
	_GDIPlus_PenDispose($hPen)

	; рисуем линию обсалютного лимита
	$hPen = _GDIPlus_PenCreate(0xFFFF0000)
    _GDIPlus_GraphicsDrawLine($hGraphic, 0, $YOffset-150, $WX + 1, $YOffset-150, $hPen)
	_GDIPlus_PenDispose($hPen)

	; рисуем линию потеря пакета
	$hPen = _GDIPlus_PenCreate(0xFFF0F0F0)
    _GDIPlus_GraphicsDrawLine($hGraphic, 0, $YOffset-255, $WX + 1, $YOffset-255, $hPen)
	_GDIPlus_PenDispose($hPen)

	$hPen = _GDIPlus_PenCreate(0xFF000000)
    For $i = 0 To UBound($aPoints) - 1
        $Xi = $XOffset + $XScale * $aPoints[$i][0]
		$Yi = $YOffset - $YScale * $aPoints[$i][1]

		;Убрать лост пакеты
		;If $Yi < 50 Then $Yi = 301

        If $i Then
            _GDIPlus_GraphicsDrawLine($hGraphic, $Xp, $Yp, $Xi, $Yi, $hPen)
        EndIf
        $Xp = $Xi
        $Yp = $Yi
    Next
    $hBitmap = _GDIPlus_BitmapCreateHBITMAPFromBitmap($hImage)
    _GDIPlus_GraphicsDispose($hGraphic)
    _GDIPlus_ImageDispose($hImage)
    _GDIPlus_BrushDispose($hBrush)
    _GDIPlus_PenDispose($hPen)
    _GDIPlus_Shutdown()

    $hForm = GUICreate('My Graph', $WX + 1, $WY + 1)
    $Pic = GUICtrlCreatePic('', 0, 0, $WX + 1, $WY + 1)
    $hPic = GUICtrlGetHandle(-1)

    _WinAPI_DeleteObject(_SendMessage($hPic, 0x0172, 0, $hBitmap))
    $hObj = _SendMessage($hPic, 0x0173)
    If $hObj <> $hBitmap Then
        _WinAPI_DeleteObject($hBitmap)
    EndIf

    GUISetState(@SW_SHOW, $hForm)

    Do
    Until GUIGetMsg() = -3

    GUIDelete($hForm)

EndFunc   ;==>_ViewGraph


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
	FileWrite($hFileOpen, "_______________________ от " & @MDAY & '.' & @MON & '.' & @YEAR & "." & @CRLF & @CRLF)
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
        ;If Not _SQLite_Exec ($dbn, "CREATE TABLE tblPing (_id integer PRIMARY KEY AUTOINCREMENT, DATETIME DEFAULT (datetime('now','localtime')), Server, Ping integer);") = $SQLITE_OK Then _
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

Func _MainProcess()
EndFunc

;Main
MainInit()
MainLoop()
MainExit()
;End Main
