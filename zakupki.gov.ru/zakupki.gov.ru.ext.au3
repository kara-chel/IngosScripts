; +TODO Общий SQL
; +TODO Общий INI
Opt("TrayIconHide", 0) ;0 - отображать, 1 - скрыть

#include <SQLite.au3>
#include <SQLite.dll.au3>
#include <File.au3>
#include <MD5.AU3>
#include <IE.au3>

Const $cVersion = "ZakupkiGovRuExtIngos20170503"
Const $cProduct = "zakupki.gov.ru.ext"
Const $cZURL = "http://zakupki.gov.ru/epz/order/extendedsearch/search.html?placeOfSearch=FZ_44&placeOfSearch=FZ_223&orderPriceFrom=&orderPriceTo=&orderPriceCurrencyId=-1&deliveryAddress=&participantName=&orderPublishDateFrom=&orderPublishDateTo=&orderUpdateDateFrom=&orderUpdateDateTo=&customer.title=&customer.code=&customer.fz94id=&customer.fz223id=&customer.inn=&agency.title=&agency.code=&agency.fz94id=&agency.fz223id=&agency.inn=&regionIds=5277380&orderStages=AF&searchTextInAttachedFile=&applSubmissionCloseDateFrom=&applSubmissionCloseDateTo=&searchString=%D1%81%D1%82%D1%80%D0%B0%D1%85&morphology=false&strictEqual=false"
Const $cHost = "http://zakupki.gov.ru"


Global $hLogFile
Global $sLOGFile = @ScriptDir & "\log\" & $cProduct & ".[" & @UserName & "].[" & @IPAddress1 & "].log"
Global $sSQLFile = @ScriptDir & "\db\zakupki.gov.ru.[" & @UserName & "].sql"
Global $sINIFile = @ScriptDir & "\zakupki.gov.ru.ini"
Global $sDLLFile = @ScriptDir & "\sqlite3.dll"
Global $sURL
Global $sZURL
Global $sHost
Global $oIE
Global $retarr, $dbn
Global $iAttach
Global $DEBUGMODE = True
Global $iIEVisible

Func MainExit()
	_FileWriteLog($hLogFile, "Скрипт завершен")
	_CloseDB(); Закрытие ДБ
	FileClose($hLogFile); Закрываем LOG-файл
	Exit
EndFunc

Func MainInit()

	$hLogFile = FileOpen($sLogFile, $FO_APPEND)
	; Проверка на повторный запуск скрипта
	If WinExists($cVersion) Then
		_FileWriteLog($hLogFile, "MainInit(): Скрипт уже запущен, выход. ")
		FileClose($hLogFile); Закрываем LOG-файл
		Exit
	EndIf
	AutoItWinSetTitle($cVersion)
	;$hExist = GUICreate($cVersion, 0, 0, 0, 0)
	;GUISetState(@SW_HIDE)
	; /

	If $hLogFile = -1 Then
		MsgBox($MB_SYSTEMMODAL, "", "An error occurred when openning the file.")
	EndIf
	_FileWriteLog($hLogFile, "Скрипт запущен")
	;TrayTip($cProduct, "Скрипт запущен", 3, 1)

	If $DEBUGMODE Then _FileWriteLog($hLogFile, "MainInit(): @ScriptDir = " & @ScriptDir)

	;***** URL *****
	$sZURL = IniRead($sINIFile, @UserName, "URL", "")
	If $sZURL = "" Then $sZURL = IniRead($sINIFile, "General", "URL", $cZURL)
	;***** HOST *****
	$sHost = IniRead($sINIFile, @UserName, "HOST", "")
	If $sHost = "" Then $sHost = IniRead($sINIFile, "General", "HOST", $cHost)
	;***** DEBUGMODE *****
	Local $sTmp = IniRead($sINIFile, @UserName, "DEBUGMODE", "")
	If $sTmp = "" Then $sTmp = IniRead($sINIFile, "General", "DEBUGMODE", "False")
	If StringUpper($sTmp) = "TRUE" Then
		$DEBUGMODE = True
	Else
		$DEBUGMODE = False
	EndIf
	;***** Visible *****
	Local $sTmp = IniRead($sINIFile, @UserName, "Visible", "")
	If $sTmp = "" Then $sTmp = IniRead($sINIFile, "General", "Visible", "False")
	If StringUpper($sTmp) = "TRUE" Then
		$iIEVisible = 1
	Else
		$iIEVisible = 0
	EndIf
	;***** Attach *****
	Local $sTmp = IniRead($sINIFile, @UserName, "Attach", "")
	If $sTmp = "" Then $sTmp = IniRead($sINIFile, "General", "Attach", "False")
	If StringUpper($sTmp) = "TRUE" Then
		$iAttach = 1
	Else
		$iAttach = 0
	EndIf

	If $DEBUGMODE Then _FileWriteLog($hLogFile, "MainInit(): $DEBUGMODE = " & $DEBUGMODE)
	If $DEBUGMODE Then _FileWriteLog($hLogFile, "MainInit(): $sZURL = " & $sZURL)
	If $DEBUGMODE Then _FileWriteLog($hLogFile, "MainInit(): $sHost = " & $sHost)
	If $DEBUGMODE Then _FileWriteLog($hLogFile, "MainInit(): $iAttach = " & $iAttach)

	_OpenDB(); Открытие/создание ДБ

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
	_MainProcess()
EndFunc

Func ChckUrl($sStr)
	If StringLeft($sStr, 1) = '/' Then
		Return $sHost & $sStr
	Else
		Return $sStr
	EndIf
EndFunc

Func _CheckRecordDB($sTitle, $sURL)
	 _SQLite_QuerySingleRow($dbn,"SELECT * FROM tblTender WHERE Title='" & $sTitle & "' AND URL='" & $sURL & "' AND UserName='" & StringUpper(@UserName) & "'", $retarr)
    If $retarr[0] <> "" Then
		Return True
	Else
		Return False
	EndIf
EndFunc

Func _AddRecordDB($sTitle, $sURL, $sDate = @MDAY & '.' & @MON & '.' & @YEAR, $sTime = @HOUR & ':' & @MIN & ':' & @SEC)
	If Not _SQLite_Exec ($dbn, "INSERT INTO tblTender(UserName, Title, URL, Date, Time) VALUES ('" & StringUpper(@UserName) & "','" & $sTitle & "','" & _
	 $sURL & "','" & $sDate & "','" & $sTime & "');") = $SQLITE_OK Then _
            MsgBox(16, "SQLite Ошибка", _SQLite_ErrMsg ())
EndFunc

Func _MainProcess()
	If $DEBUGMODE Then _FileWriteLog($hLogFile, "_MainProcess(): Enter")
	Local $oIE = _IECreate ( "about:blank", 0, $iIEVisible, 1)
	Local $oTD
	Local $oA
	Local $sTitle, $sURL
	Local $oIETender

	; Делаем ссылку уникальной добавляем вымышленный параметр orderPriceCurrencyKey
	$sSuffix = "&orderPriceCurrencyKey=" & md5(@UserName & @YEAR & @MON & @MDAY & @HOUR & @MIN & @SEC)
	If $DEBUGMODE Then _FileWriteLog($hLogFile, "_MainProcess(): URL -  " & $sZURL & $sSuffix)

	If $DEBUGMODE Then _FileWriteLog($hLogFile, "_MainProcess(): _IENavigate")
	_IENavigate ($oIE, $sZURL & $sSuffix, 1)
	If $DEBUGMODE Then _FileWriteLog($hLogFile, "_MainProcess(): _IETagNameGetCollection($oIE, 'td')")
	Local $oTDs = _IETagNameGetCollection($oIE, "td")
	;Local $oTDs = $oIE.document.GetElementsByTagName("td")
	If Not @error Then
		For $oTD In $oTDs
			If $oTD.ClassName == "descriptTenderTd" Then
				If $DEBUGMODE Then _FileWriteLog($hLogFile, "_MainProcess(): _IETagNameGetCollection($oTD, 'a')")
				Local $oAs = _IETagNameGetCollection($oTD, "a")
				;Local $oAs = $oTD.document.GetElementsByTagName("a")
				If Not @error Then
					For $oA in $oAs
						; проверить есть ли запись в базе
						; если есть то ничего не делать
						; если нет, то открыть окно IE и добавить запись
						$sTitle = StringStripWS($oA.innertext, 1 + 2 + 4 + 8)
						$sURL = ChckUrl($oA.getAttributeNode('href').nodeValue)
						If $DEBUGMODE Then _FileWriteLog($hLogFile, "_MainProcess(): $sTitle = " & $sTitle)
						If $DEBUGMODE Then _FileWriteLog($hLogFile, "_MainProcess(): $sURL = " & $sURL)
						If Not _CheckRecordDB($sTitle, $sURL) Then
							_FileWriteLog($hLogFile, "Новый тендер: " & $sTitle & " - " & $sURL)
							_AddRecordDB($sTitle, $sURL); Добавить запись
							if $iAttach Then
								Local $i = 1, $oIETmp, $bNew = True
								While 1
									If $DEBUGMODE Then _FileWriteLog($hLogFile, "_MainProcess(): $i=" & $i)
									$oIETmp = _IEAttach("", "instance", $i)
									If _IEPropertyGet($oIETmp, "hwnd") <> _IEPropertyGet($oIE, "hwnd") Then
										If $DEBUGMODE Then _FileWriteLog($hLogFile, "_MainProcess(): $oIETmp.hwnd=" & _IEPropertyGet($oIETmp, "hwnd"))
										If _IEPropertyGet($oIETmp, "hwnd") = 0 Then
											$bNew = True
										Else
											$bNew = False
										EndIf
										ExitLoop
									EndIf
									If @error = $_IEStatus_NoMatch Then
										If $DEBUGMODE Then _FileWriteLog($hLogFile, "_MainProcess(): @error=" & $_IEStatus_NoMatch)
										$bNew = True
										ExitLoop
									EndIf
									$i += 1
								WEnd
								If $bNew Then
									If $DEBUGMODE Then _FileWriteLog($hLogFile, "_MainProcess(): $bNew=" & "True")
									$oIETender = _IECreate($sURL, 0, 1, 0) ;открыть окно с тендером
									Sleep(1000)
								Else
									If $DEBUGMODE Then _FileWriteLog($hLogFile, "_MainProcess(): $bNew=" & "False")
									__IENavigate($oIETmp, $sURL, 0, 0x800)
								EndIf
							Else
								$oIETender = _IECreate($sURL, 0, 1, 0) ;открыть окно с тендером
							EndIf
							Sleep(100)
						EndIf
						ExitLoop
					Next
				EndIf
			EndIf
		Next
	EndIf
	_IEQuit($oIE); Закрываем IE
	If $DEBUGMODE Then _FileWriteLog($hLogFile, "_MainProcess(): End")
EndFunc

;Main
MainInit()
MainLoop()
MainExit()
;End Main
