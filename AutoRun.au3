;TODO + INI файл с настройками
;TODO - Добавить исключение
Opt("TrayIconHide", 1) ;0 - отображать, 1 - скрыть
Opt("TrayAutoPause", 0)
Opt("TrayMenuMode", 3) ; The default tray menu items will not be shown and items are not checked when selected. These are options 1 and 2 for TrayMenuMode.
Opt("TrayOnEventMode", 1) ; Enable TrayOnEventMode.

#include <WinAPIFiles.au3>
#include <File.au3>

Const $cTime = 60 * 1000 ;sec
Const $cStartTime = @MDAY & '.' & @MON & '.' & @YEAR & ' ' & @HOUR & ':' & @MIN & ':' & @SEC
Const $cVersion = "AutoRunIngos20160310"
Global $sScriptsDir
Global $tTimer, $tSleep
Global $hFile
Global $idMainInfo, $idMainQuit
Global $hExist
Global $aAdmin

Func MainCheckCron($bStartUp = False)

	Local $aLines = FileReadToArray($sScriptsDir & "/AutoRun.cron")
	If @error Then
		;MsgBox($MB_SYSTEMMODAL, "Ошибка!", "Ошибка открытия CRON-файла, обратитесь к ИТ специалисту! @error: " & @error)
		Sleep(1000)
	Else
		For $i = 0 To UBound($aLines) - 1 ; Loop through the array.
			$aLines[$i] = StringUpper($aLines[$i])
			; Убрать комментарии
			Local $iPosition = StringInStr($aLines[$i], ";")
			if $iPosition > 0 Then
				$aLines[$i] = StringLeft($aLines[$i], $iPosition - 1)
				;MsgBox($MB_SYSTEMMODAL, "", $aLines[$i])
			EndIf
			; Убрать пробелы 3 - в начале и в конце; 8 - все
			$aLines[$i] = StringStripWS($aLines[$i], 3)
			; проверяем чтоб было 4 параметра
			Local $aColumn = StringSplit($aLines[$i], "|")
			If $aColumn[0] <> 4 Then ContinueLoop

			; Проверить первый параметр - IP, UserName, *
			If $aColumn[1] <> '*' And $aColumn[1] <> @IPAddress1 And $aColumn[1] <> StringUpper(@UserName) Then ContinueLoop
			; Проверить второй параметр - день
			; Вместо DAY можно указать числовой день недели. Диапазон составляет от 0 до 6, что соответствует воскресенья=0 по субботу=6.
			; @MDAY.@MON.@YEAR ;*.@MON.@YEAR ;@MDAY.*.@YEAR ;@MDAY.@MON.*
			If $aColumn[2] <> '*' And $aColumn[2] <> @MDAY & '.' & @MON & '.' & @YEAR And $aColumn[2] <> @WDAY - 1 _
				And $aColumn[2] <> '*.' & @MON & '.' & @YEAR And $aColumn[2] <> @MDAY & '.*.' & @YEAR And $aColumn[2] <> @MDAY & '.' & @MON & '.*' _
				And $aColumn[2] <> '*.*.' & @YEAR And $aColumn[2] <> '*.' & @MON & '.*' And $aColumn[2] <> @MDAY & '.*.*' Then ContinueLoop
			; Проверить третий параметр - время ;*:@MIN ; @HOUR:* ; *
			If ($aColumn[3] <> 'STARTUP' Or $bStartUp <> True) And $aColumn[3] <> @HOUR & ':' & @MIN _
				And $aColumn[3] <> '*:' & @MIN And $aColumn[3] <> @HOUR & ':*' And $aColumn[3] <> '*' Then ContinueLoop
			; MsgBox($MB_SYSTEMMODAL, "Запускаем", $aLines[$i])
			If FileExists($sScriptsDir & "\AutoRun\" & $aColumn[4]) Then
				_FileWriteLog($hFile, "Cron [" & $aLines[$i] & "] - ShellExecute: " & $sScriptsDir & "\AutoRun\" & $aColumn[4])
				ShellExecute($sScriptsDir & "\AutoRun\" & $aColumn[4])
			Else
				_FileWriteLog($hFile, "ErrorCron [" & $aLines[$i] & "] - FileExists:" & $sScriptsDir & "\AutoRun\" & $aColumn[4])
			EndIf
        Next
    EndIf
EndFunc

Func MainInfo()
	Local $sResult = "Время запуска скрипта: " & $cStartTime & @CRLF & _
		@CRLF & _
		"Имя компьютера: " & StringUpper(@ComputerName) & @CRLF & _
		"Имя пользователя: " & StringUpper(@UserName) & @CRLF & _
		@CRLF & _
		"IP адрес: " & @IPAddress1
	Local Const $sFilePath = _WinAPI_GetTempFileName(@TempDir)

	Local $hFileOpen = FileOpen($sFilePath, $FO_APPEND)
	If $hFileOpen = -1 Then
		MsgBox($MB_SYSTEMMODAL, $cVersion, "Ошибка открытия временного файла", 5)
		;Return False
	EndIf
	FileWrite($hFileOpen, $sResult)
	FileClose($hFileOpen)

	Run("notepad.exe " & $sFilePath)
	WinWaitActive("[CLASS:Notepad]")
	Sleep(500)

	FileDelete($sFilePath)

EndFunc

Func MainStartUp()
	MainCheckCron(True)
EndFunc

Func MainInit()

	; Проверка на повторный запуск скрипта
	If WinExists($cVersion) Then Exit
	AutoItWinSetTitle($cVersion)
	;$hExist = GUICreate($cVersion, 0, 0, 0, 0)
	;GUISetState(@SW_HIDE)
	; /

	; загрузить параметры с AutoRun.ini
	$aAdmin = StringSplit(StringUpper(StringStripWS(IniRead(@ScriptDir & "/AutoRun.ini", "General", "Admin", ''), 8)), ',')
	$sScriptsDir = StringStripWS(IniRead(@ScriptDir & "/AutoRun.ini", "General", "ScriptsDir", @ScriptDir), 8)
	; /

	; Логи
	$hFile = FileOpen($sScriptsDir & "\Log\" & "[" & @UserName & "][" & @ComputerName & "][" & @IPAddress1 & "]" & ".log", 1)
	_FileWriteLog($hFile, "@UserName: " & @UserName)
	_FileWriteLog($hFile, "@ComputerName: " & @ComputerName)
	_FileWriteLog($hFile, "@LogonDomain: " & @LogonDomain)
	_FileWriteLog($hFile, "@IPAddress1: " & @IPAddress1)
	; /

	; Таймер
	$tTimer = TimerInit()
	$tSleep = $cTime
	; /

	; Трэй
	$idMainInfo = TrayCreateItem("Info")
	TrayItemSetOnEvent(-1, "MainInfo")
	TrayCreateItem("")
	$idMainQuit = TrayCreateItem("Exit")
	TrayItemSetOnEvent(-1, "MainQuit")
	TraySetState()
	TraySetIcon("Shell32.dll", -13)
	; /

EndFunc

Func MainLoop()
	MainStartUp()
	While 1
		if TimerDiff($tTimer) < $tSleep Then
			sleep(100)
		Else
			MainCheckCron()
			$tSleep = TimerDiff($tTimer) + $cTime
		EndIf
	WEnd
EndFunc

Func MainQuit()
	; Только указанные пользователи могут выйти.
    Local $iFind = _ArraySearch($aAdmin, StringUpper(@UserName))
	if  $iFind >= 0 Then
		GUIDelete($hExist)
		FileClose($hFile)
		Exit
	EndIf
EndFunc

MainInit()
MainLoop()
MainQuit()
