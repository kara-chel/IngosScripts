Opt("TrayIconHide", 0) ;0 - ����������, 1 - ������

#include <SQLite.au3>
#include <SQLite.dll.au3>
#include <File.au3>
#include <MD5.AU3>
#include <IE.au3>

Const $cVersion = "RtsTenderRuExtIngos20190918"
Const $cProduct = "rts-tender.ru.ext"
Const $cZURL = "https://www.rts-tender.ru/auctionsearch?purches_name=%D1%81%D1%82%D1%80%D0%B0%D1%85"
Const $cHost = "https://www.rts-tender.ru"


Global $hLogFile
Global $sLOGFile = @ScriptDir & "\log\" & $cProduct & ".[" & @UserName & "].[" & @IPAddress1 & "].log"
Global $sSQLFile = @ScriptDir & "\db\rts-tender.ru.[" & @UserName & "].sql"
Global $sINIFile = @ScriptDir & "\rts-tender.ru.ini"
Global $sDLLFile = @ScriptDir & "\sqlite3.dll"
Global $sURL
Global $sZURL
Global $sHost
Global $oIE
Global $retarr, $dbn
Global $iAttach
Global $DEBUGMODE = True
Global $iIEVisible = True

Func MainExit()
	_FileWriteLog($hLogFile, "������ ��������")
	_CloseDB(); �������� ��
	FileClose($hLogFile); ��������� LOG-����
	Exit
EndFunc

Func MainInit()

	$hLogFile = FileOpen($sLogFile, $FO_APPEND)
	; �������� �� ��������� ������ �������
	If WinExists($cVersion) Then
		_FileWriteLog($hLogFile, "MainInit(): ������ ��� �������, �����. ")
		FileClose($hLogFile); ��������� LOG-����
		Exit
	EndIf
	$hExist = GUICreate($cVersion, 0, 0, 0, 0)
	GUISetState(@SW_HIDE)
	; /

	If $hLogFile = -1 Then
		MsgBox($MB_SYSTEMMODAL, "", "An error occurred when openning the file.")
	EndIf
	_FileWriteLog($hLogFile, "������ �������")
	;TrayTip($cProduct, "������ �������", 3, 1)

	If $DEBUGMODE Then _FileWriteLog($hLogFile, "MainInit(): @ScriptDir = " & @ScriptDir)

	;***** URL *****
	$sZURL = IniRead($sINIFile, @UserName, "URL", "")
	If $sZURL = "" Then $sZURL = IniRead($sINIFile, "General", "URL", $cZURL)
	;***** HOST *****
	$sHost = IniRead($sINIFile, @UserName, "HOST", "")
	If $sHost = "" Then $sHost = IniRead($sINIFile, "General", "HOST", $cHost)
	;***** DEBUGMODE *****
	Local $sTmp = IniRead($sINIFile, @UserName, "DEBUGMODE", "")
	If $sTmp = "" Then $sTmp = IniRead($sINIFile, "General", "DEBUGMODE", "True")
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
	If $DEBUGMODE Then _FileWriteLog($hLogFile, "MainInit(): $iIEVisible = " & $iIEVisible)
	If $DEBUGMODE Then _FileWriteLog($hLogFile, "MainInit(): $iAttach = " & $iAttach)

	_OpenDB(); ��������/�������� ��

EndFunc

Func _OpenDB()
	If @AutoItX64 And (StringInStr($sDLLFile, "_x64") = 0) Then $sDLLFile = StringReplace($sDLLFile, ".dll", "_x64.dll")
	If $DEBUGMODE Then _FileWriteLog($hLogFile, "_OpenDB(): $sDLLFile = " & $sDLLFile)
	_SQLite_Startup ($sDLLFile, False, 1)
    If @error > 0 Then
        MsgBox(16, "SQLite ������", "DLL �� ����� ���� ��������!")
        Exit - 1
    EndIf
    If NOT FileExists($sSQLFile) Then
        $dbn=_SQLite_Open($sSQLFile)
        If @error > 0 Then
            MsgBox(16, "SQLite ������", "�� �������� ������� ����!")
            Exit - 1
        EndIf
        If Not _SQLite_Exec ($dbn, "CREATE TABLE tblTender (_id integer PRIMARY KEY AUTOINCREMENT, UserName, Title, URL, Date, Time);") = $SQLITE_OK Then _
            MsgBox(16, "SQLite ������", _SQLite_ErrMsg ())
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
            MsgBox(16, "SQLite ������", _SQLite_ErrMsg ())
EndFunc

Func _MainProcess()
	If $DEBUGMODE Then _FileWriteLog($hLogFile, "_MainProcess(): Enter")
	If $DEBUGMODE Then ConsoleWrite("*** _MainProcess(): $iIEVisible=" & $iIEVisible & @CRLF)
	Local $oIE = _IECreate ( "about:blank", 0, $iIEVisible, 1)
	Local $oTD
	Local $oA
	Local $sTitle, $sURL
	Local $oIETender

	; ������ ������ ���������� ��������� ����������� �������� orderPriceCurrencyKey
	$sSuffix = "";"&orderPriceCurrencyKey=" & md5(@UserName & @YEAR & @MON & @MDAY & @HOUR & @MIN & @SEC)
	If $DEBUGMODE Then _FileWriteLog($hLogFile, "_MainProcess(): URL -  " & $sZURL & $sSuffix)

	If $DEBUGMODE Then _FileWriteLog($hLogFile, "_MainProcess(): _IENavigate")
	_IENavigate ($oIE, $sZURL & $sSuffix, 1)
	If @error Then _FileWriteLog($hLogFile, "_MainProcess(): _IENavigate@error=" & @error)
	_IELoadWait($oIE) ;���� �� ���� �����, �� ��� ����� ���� ���������� ��������

	Local $oForm = _IEFormGetObjByName($oIE, "Form")
	If @error Then _FileWriteLog($hLogFile, "_MainProcess(): Form@error=" & @error)
	; dnn$ctr691$View$aSclassRegion: ����������� �������
	Local $oRegion = _IEFormElementGetObjByName($oForm, "dnn$ctr691$View$aSclassRegion")
	_IEFormElementSetValue($oRegion, "����������� �������")
	If @error Then _FileWriteLog($hLogFile, "_MainProcess(): dnn$ctr691$View$aSclassRegion@error=" & @error)
	; dnn$ctr691$View$aSclassRegionCode: 7400000000000
	Local $oRegionCode = _IEFormElementGetObjByName($oForm, "dnn$ctr691$View$aSclassRegionCode")
	_IEFormElementSetValue($oRegionCode, "7400000000000")
	If @error Then _FileWriteLog($hLogFile, "_MainProcess(): dnn$ctr691$View$aSclassRegionCode@error=" & @error)
	; dnn$ctr691$View$aSPurchesName: �������
	Local $oName = _IEFormElementGetObjByName($oForm, "dnn$ctr691$View$aSPurchesName")
	_IEFormElementSetValue($oName, "�������")
	If @error Then _FileWriteLog($hLogFile, "_MainProcess(): dnn$ctr691$View$aSPurchesName@error=" & @error)
	; dnn$ctr691$View$aScomplaintSelect: No
	Local $oComplaintSelect = _IEFormElementGetObjByName($oForm, "dnn$ctr691$View$aScomplaintSelect")
	_IEFormElementSetValue($oComplaintSelect, "No")
	If @error Then _FileWriteLog($hLogFile, "_MainProcess(): dnn$ctr691$View$aScomplaintSelect@error=" & @error)
	; dnn$ctr691$View$aSbuttonSearch: �����
	Local $oButtonSearch = _IEFormElementGetObjByName($oForm, "dnn$ctr691$View$aSbuttonSearch")
	_IEFormElementSetValue($oButtonSearch, "�����")
	If @error Then _FileWriteLog($hLogFile, "_MainProcess(): dnn$ctr691$View$aSbuttonSearch@error=" & @error)
	; dnn$ctr691$View$GroupFilter: RadioButton1
	Local $oGroupFilter = _IEFormElementGetObjByName($oForm, "dnn$ctr691$View$GroupFilter")
	_IEFormElementSetValue($oGroupFilter, "�����")
	If @error Then _FileWriteLog($hLogFile, "_MainProcess(): dnn$ctr691$View$GroupFilter@error=" & @error)

	; Submit
	;Sleep(4000)
	;_IEFormSubmit($oForm);�� ��������� �������� ����
	;dnn$ctr691$View$aSbuttonSearch
	$oBtn = _IEGetObjByName ($oIE, "dnn$ctr691$View$aSbuttonSearch")
	_IEAction($oBtn, "click")

	;������� �������� ��������
	;Sleep(2000)
	_IELoadWait($oIE)


	;��������� ����������� ��������
	;����� �������
	;<table class="table table4" id="dnn_ctr691_View_procResultList">
	;If $DEBUGMODE Then ConsoleWrite("*** Local $oTables = _IETagNameGetCollection($oIE, table) ***" & @CRLF)
	Local $oTables = _IETagNameGetCollection($oIE, "table")
	If Not @error Then
		For $oTable In $oTables
			If $oTable.ClassName == "table table4" Then

				Local $oTRs = _IETagNameGetCollection($oTable, "tr")
				If Not @error Then
					Local $numDR = 0
					For $oTR In $oTRs
						; If $numDR > 0 Then
						Local $oTDs = _IETagNameGetCollection($oTR, "td")
						If Not @error Then
							Local $numDT = 0
							For $oTD In $oTDs
								Switch $numDT
									Case 0
										;��
									Case 1
										;���� � ����� ����������
									Case 2
										;����� ���������
										Local $oAs = _IETagNameGetCollection($oTD, "a")
										If Not @error Then
											For $oA in $oAs

												$sTitle = StringStripWS($oA.innertext, 1 + 2 + 4 + 8)
												$sURL = ChckUrl($oA.getAttributeNode('href').nodeValue)
												; ��������� ���� �� ������ � ����
												; ���� ���� �� ������ �� ������
												; ���� ���, �� ������� ���� IE � �������� ������
												If Not _CheckRecordDB($sTitle, $sURL) Then
													_FileWriteLog($hLogFile, "����� ������: " & $sTitle & " - " & $sURL)
													If $DEBUGMODE Then _FileWriteLog($hLogFile, "_MainProcess(): _AddRecordDB=" & $sTitle)
													_AddRecordDB($sTitle, $sURL); �������� ������ EndSwitch
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
															$oIETender = _IECreate($sURL, 0, 1, 0) ;������� ���� � ��������
															Sleep(1000)
														Else
															If $DEBUGMODE Then _FileWriteLog($hLogFile, "_MainProcess(): $bNew=" & "False")
															__IENavigate($oIETmp, $sURL, 0, 0x800)
														EndIf
													Else
														$oIETender = _IECreate($sURL, 0, 1, 0) ;������� ���� � ��������
													EndIf
													Sleep(100)
												EndIf
											Next
										EndIf
									Case 3
										;����������� �������
									Case 4
										;������ ��������
									Case 5
										;������������ ������� �������
									Case 6
										;��������� ����
									Case 7
										;������
									Case 8
										;���� ������
									Case 9
										;���������� ������
									Case 10
										;������ ������
									Case 11
										;���
									Case 12
										;������
									Case 13
										;������



								EndSwitch
								$numDT = $numDT + 1
							Next
						EndIf
						$numDR = $numDR + 1

					Next
				EndIf

			EndIf
		Next
	EndIf

	;If Not $DEBUGMODE Then
	_IEQuit($oIE); ��������� IE
	If $DEBUGMODE Then _FileWriteLog($hLogFile, "_MainProcess(): End")
EndFunc

;Main
MainInit()
MainLoop()
MainExit()
;End Main
