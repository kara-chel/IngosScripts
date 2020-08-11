#include-once
#include <Date.au3> ; _FO_FileBackup


; FileOperations.au3
; ������ 1.5 �� 2012.08.08

; http://www.autoitscript.com/forum/topic/133224-filesearch-foldersearch/
; http://pastebin.com/AbzMyg1x
; http://azjio.ucoz.ru/publ/skripty_autoit3/funkcii/filesearch/11-1-0-33

; =======================================
; Title .........: FileOperations
; AutoIt Version : 3.3.2.0 - 3.3.8.1
; Language ......: English + �������
; Description ...: Operations with files
; =======================================

; #CURRENT# =============================
; _FO_CorrectMask
; _FO_FileSearch
; _FO_FolderSearch
; _FO_SearchEmptyFolders
; _FO_FileDirReName
; _FO_GetCopyName
; _FO_FileBackup
; _FO_PathSplit
; =======================================

; ���������� �������
; #INTERNAL_USE_ONLY#====================
; __FO_FileSearchType
; __FO_FileSearchMask
; __FO_FileSearchAll
; __FO_GetListMask
; __FO_MaskUnique (#Obfuscator_Off � #Obfuscator_On)
; __FO_FolderSearch
; __FO_FolderSearchMask
; __FO_SearchEmptyFolders1
; __FO_SearchEmptyFolders2
; __FO_UserLocale
; =======================================

; #FUNCTION# ;=================================================================================
; Function Name ...: _FO_FileSearch (__FO_FileSearchType, __FO_FileSearchMask, __FO_FileSearchAll)
; AutoIt Version ....: 3.3.2.0+ , versions below this @extended should be replaced by of StringInStr(FileGetAttrib($sPath&'\'&$sFile), "D")
; Description ........: Search files by mask in subdirectories.
; Syntax................: _FO_FileSearch ( $sPath [, $sMask = '*' [, $fInclude=True [, $iDepth=125 [, $iFull=1 [, $iArray=1 [, $iTypeMask=1 [, $sLocale=0]]]]]]] )
; Parameters:
;		$sPath - search path
;		$sMask - allowed two options for the mask: using symbols "*" and "?" with the separator "|", or a list of extensions with the separator "|"
;		$fInclude - (True / False) invert the mask, that is excluded from the search for these types of files
;		$iDepth - (0-125) nesting level (0 - root directory)
;		$iFull - (0,1,2,3)
;                  |0 - relative
;                  |1 - full path
;                  |2 - file names with extension
;                  |3 - file names without extension
;		$iArray - if the value other than zero, the result is an array (by default ),
;                  |0 - a list of paths separated by @CRLF
;                  |1 - array, where $iArray[0]=number of files ( by default)
;                  |2 - array, where $iArray[0] contains the first file
;		$iTypeMask - (0,1,2) defines the format mask
;                  |0 - auto detect
;                  |1 - forced mask, for example *.is?|s*.cp* (it is possible to specify a file name with no characters * or ? and no extension will be found)
;                  |2 - forced mask, for example tmp|bak|gid (that is, only files with the specified extension)
;		$sLocale - case sensitive.
;                  |0 - not case sensitive (only for 'A-z'), by default.
;                  |1 - case sensitive (for any characters)
;                  |<symbols> - not case sensitive, specified range of characters from local languages. For example '�-���'. 'A-z' is not required, they are enabled by default.
; Return values ....: Success - Array or a list of paths separated by @CRLF
;					Failure - empty string, @error:
;                  |0 - no error
;                  |1 - Invalid path
;                  |2 - Invalid mask
;                  |3 - not found
; Author(s) ..........: AZJIO
; Remarks ..........: Use function _CorrectMask if it is required correct mask, which is entered by user
; ============================================================================================
; ��� ������� ...: _FO_FileSearch (__FO_FileSearchType, __FO_FileSearchMask, __FO_FileSearchAll)
; ������ AutoIt ..: 3.3.2.0+ , � ������� ���� ��������� ����� @extended �������� �� StringInStr(FileGetAttrib($sPath&'\'&$sFile), "D")
; �������� ........: ����� ������ �� ����� � ������������.
; ���������.......: _FO_FileSearch ( $sPath [, $sMask = '*' [, $fInclude=True [, $iDepth=125 [, $iFull=1 [, $iArray=1 [, $iTypeMask=1 [, $sLocale=0]]]]]]] )
; ���������:
;		$sPath - ���� ������
;		$sMask - ��������� ��� �������� �����: � �������������� �������� "*" � "?" � ������������� ����� "|", ���� ������������ ���������� ����� "|"
;		$fInclude - (True / False) ������������� �����, �� ���� ��������� �� ������ ��������� ���� ������
;		$iDepth - (0-125) ������� ����������� (0 - �������� �������)
;		$iFull - (0,1,2,3)
;                  |0 - �������������
;                  |1 - ������ ����
;                  |2 - ����� ������ � �����������
;                  |3 - ����� ������ ��� ����������
;		$iArray - (0,1,2) ���������� ����� ����������, ������ ��� ������
;                  |0 - ������ � ������������ @CRLF
;                  |1 - ������, � ������� $iArray[0]=���������� ������ (�� ���������)
;                  |2 - ������, � ������� $iArray[0] �������� ������ ����
;		$iTypeMask - (0,1,2) ���������� ������ ������ �����
;                  |0 - ���������������
;                  |1 - ������������� ����� ���� *.is?|s*.cp* (�� ���� ����� ������� ��� ����� ��� �������� * ��� ? � ��� ���������� � ����� �������)
;                  |2 - ������������� ����� ���� tmp|bak|gid (�� ����������, �� ���� ������ ����� ������ � ��������� �����������)
;		$sLocale - ��������� ������� ���� ��� ������.
;                  |0 - �� ��������� ������� (������ ��� ��������� ����), �� ���������.
;                  |1 - ��������� ������� (��� ����� ��������)
;                  |<�������> - �� ��������� ������� ���������� ��������� �������� ���������� �����, ������� ���������, �������� '�-���'. ��������� ��������� �� ���������, ��� �� ��������� ��������.
; ������������ ��������: ������� - ������ ��� ������ � ������������ @CRLF
;					�������� - ������ ������, @error:
;                  |0 - ��� ������
;                  |1 - �������� ����
;                  |2 - �������� �����
;                  |3 - ������ �� �������
; ����� ..........: AZJIO
; ���������� ..: ����������� ������� _CorrectMask, ���� ����� ����������� �� ���� ����� � ��������� �������� �� ������������
; ============================================================================================
; ������� �������� � ���������� ������� ���������� � ��������� ��������������� ������
Func _FO_FileSearch($sPath, $sMask = '*', $fInclude = True, $iDepth = 125, $iFull = 1, $iArray = 1, $iTypeMask = 1, $sLocale = 0)
	Local $vFileList
	If $sMask = '|' Then Return SetError(2, 0, '')
	If StringRight($sPath, 1) <> '\' Then $sPath &= '\'
	; If Not StringRegExp($sPath, '(?i)^[a-z]:[^/:*?"<>|]*$') Or StringInStr($sPath, '\\') Then Return SetError(1, 0, '')
	If Not FileExists($sPath) Then Return SetError(1, 0, '')

	If $sMask = '*' Or $sMask = '' Then
		__FO_FileSearchAll($vFileList, $sPath, $iDepth)
		$vFileList = StringTrimRight($vFileList, 2)
	Else
		Switch $iTypeMask
			Case 0
				If StringInStr($sMask, '*') Or StringInStr($sMask, '?') Or StringInStr($sMask, '.') Then
					__FO_GetListMask($sPath, $sMask, $fInclude, $iDepth, $vFileList, $sLocale)
				Else
					__FO_FileSearchType($vFileList, $sPath, '|' & $sMask & '|', $fInclude, $iDepth)
					$vFileList = StringTrimRight($vFileList, 2)
				EndIf
			Case 1
				__FO_GetListMask($sPath, $sMask, $fInclude, $iDepth, $vFileList, $sLocale)
			Case Else
				If StringInStr($sMask, '*') Or StringInStr($sMask, '?') Or StringInStr($sMask, '.') Then Return SetError(2, 0, '')
				__FO_FileSearchType($vFileList, $sPath, '|' & $sMask & '|', $fInclude, $iDepth)
				$vFileList = StringTrimRight($vFileList, 2)
		EndSwitch
	EndIf

	If Not $vFileList Then Return SetError(3, 0, '')
	Switch $iFull
		Case 0
			$vFileList = StringRegExpReplace($vFileList, '(?m)^(?:.{' & StringLen($sPath) & '})(.*)$', '\1')
		Case 2
			$vFileList = StringRegExpReplace($vFileList, '(?m)^(?:.*\\)(.*)$', '\1')
		Case 3
			$vFileList = StringRegExpReplace($vFileList, '(?m)^(?:.*\\)([^\\]*?)(?:\.[^.]+)?$', '\1' & @CR)
			$vFileList = StringTrimRight($vFileList, 1)
	EndSwitch
	Switch $iArray
		Case 1
			$vFileList = StringSplit($vFileList, @CRLF, 1)
			; If @error And $vFileList[1] = '' Then Dim $vFileList[1] = [0]
		Case 2
			$vFileList = StringSplit($vFileList, @CRLF, 3)
			; If @error And $vFileList[0]='' Then SetError(3, 0, '')
	EndSwitch
	Return $vFileList
EndFunc   ;==>_FO_FileSearch

; ��������� ������ � ��������� ���������� ����������
Func __FO_GetListMask($sPath, $sMask, $fInclude, $iDepth, ByRef $sFileList, $sLocale)
	Local $aFileList, $i, $rgex
	__FO_FileSearchMask($sFileList, $sPath, $iDepth)
	$sFileList = StringTrimRight($sFileList, 2)
	$sMask = StringReplace(StringReplace(StringRegExpReplace($sMask, '[][$^.{}()+]', '\\$0'), '?', '.'), '*', '.*?')

	If Not $sLocale Then
		$rgex = 'i'
	ElseIf Not($sLocale == '1') Then
		$rgex = 'i'
		$sMask = __FO_UserLocale($sMask, $sLocale)
	EndIf

	If $fInclude Then
		$aFileList = StringRegExp($sFileList, '(?m' & $rgex & ')^(.+\|(?:' & $sMask & '))(?:\r|\z)', 3)
		$sFileList = ''
		For $i = 0 To UBound($aFileList) - 1
			$sFileList &= $aFileList[$i] & @CRLF
		Next
	Else
		$sFileList = StringRegExpReplace($sFileList & @CRLF, '(?m' & $rgex & ')^.+\|(' & $sMask & ')\r\n', '')
	EndIf
	$sFileList = StringReplace(StringTrimRight($sFileList, 2), '|', '')
EndFunc   ;==>__FO_GetListMask

Func __FO_UserLocale($sMask, $sLocale)
	Local $i, $s, $tmp
	$sLocale=StringRegExpReplace($sMask,'[^'&$sLocale&']', '') ; �������� ������ � ���������
	$tmp = StringLen($sLocale)
	For $i = 1 To $tmp
		$s = StringMid($sLocale, $i, 1)
		If $s Then ; ���� ������ ������, ������ ����� ����������
			If StringInStr($sLocale, $s, 0, 2, $i) Then
				$sLocale = $s&StringReplace($sLocale, $s, '') ; ���� ������ �����������, �� ������� ���, ��������� ��� � ������ ������
			EndIf
		Else
			ExitLoop
		EndIf
	Next
	If $sLocale Then
		$tmp = StringSplit($sLocale, '')
		For $i = 1 To $tmp[0]
			$sMask = StringReplace($sMask, $tmp[$i], '['&StringUpper($tmp[$i])&StringLower($tmp[$i])&']')
		Next
	EndIf
	Return $sMask
EndFunc   ;==>__FO_UserLocale

; ����� ��������� ����� ������
Func __FO_FileSearchType(ByRef $sFileList, $sPath, $sMask, $fInclude, $iDepth, $iCurD = 0)
	Local $iPos, $sFile, $s = FileFindFirstFile($sPath & '*')
	If $s = -1 Then Return
	While 1
		$sFile = FileFindNextFile($s)
		If @error Then ExitLoop
		If @extended Then
			If $iCurD >= $iDepth Then ContinueLoop
			__FO_FileSearchType($sFileList, $sPath & $sFile & '\', $sMask, $fInclude, $iDepth, $iCurD + 1)
		Else
			$iPos = StringInStr($sFile, ".", 0, -1)
			If $iPos And StringInStr($sMask, '|' & StringTrimLeft($sFile, $iPos) & '|') = $fInclude Then
				$sFileList &= $sPath & $sFile & @CRLF
			ElseIf Not $iPos And Not $fInclude Then
				$sFileList &= $sPath & $sFile & @CRLF
			EndIf
		EndIf
	WEnd
	FileClose($s)
EndFunc   ;==>__FO_FileSearchType

; ����� ������ �� �����
Func __FO_FileSearchMask(ByRef $sFileList, $sPath, $iDepth, $iCurD = 0)
	Local $sFile, $s = FileFindFirstFile($sPath & '*')
	If $s = -1 Then Return
	While 1
		$sFile = FileFindNextFile($s)
		If @error Then ExitLoop
		If @extended Then
			If $iCurD >= $iDepth Then ContinueLoop
			__FO_FileSearchMask($sFileList, $sPath & $sFile & '\', $iDepth, $iCurD + 1)
		Else
			$sFileList &= $sPath & '|' & $sFile & @CRLF
		EndIf
	WEnd
	FileClose($s)
EndFunc   ;==>__FO_FileSearchMask

; ����� ���� ������
Func __FO_FileSearchAll(ByRef $sFileList, $sPath, $iDepth, $iCurD = 0)
	Local $sFile, $s = FileFindFirstFile($sPath & '*')
	If $s = -1 Then Return
	While 1
		$sFile = FileFindNextFile($s)
		If @error Then ExitLoop
		If @extended Then
			If $iCurD >= $iDepth Then ContinueLoop
			__FO_FileSearchAll($sFileList, $sPath & $sFile & '\', $iDepth, $iCurD + 1)
		Else
			$sFileList &= $sPath & $sFile & @CRLF
		EndIf
	WEnd
	FileClose($s)
EndFunc   ;==>__FO_FileSearchAll


; #FUNCTION# ;=================================================================================
; Function Name ...: _FO_CorrectMask (__FO_MaskUnique)
; AutoIt Version ....: 3.3.0.0+
; Description ........: Corrects a mask
; Syntax................: _CorrectMask ( $sMask )
; Parameters........: $sMask - except symbol possible in names are allowed symbols of the substitution "*" and "?" and separator "|"
; Return values:
;					|Success -  Returns a string of a correct mask
;					|Failure - Returns a symbol "|" and @error=2
; Author(s) ..........: AZJIO
; Remarks ..........: Function corrects possible errors entered by the user
; ============================================================================================
; ��� ������� ...: _FO_CorrectMask (__FO_MaskUnique)
; ������ AutoIt ..: 3.3.0.0+
; �������� ........: ������������� �����
; ���������.......: _CorrectMask ( $sMask )
; ���������.....: $sMask - ����� �������� ���������� � ������ ����������� ������� ����������� "*" � "?" � ����������� "|"
; ������������ ��������:
;					|������� -  ���������� ������ ���������� �����
;					|�������� - ���������� ������ "|" � @error=2
; ����� ..........: AZJIO
; ���������� ..: ������� ���������� ��������� ������ ����� �������������:
; ������� ������� � ����� �� ����� ������� �������� �����, ������� ������� �������� � �����������.
; ============================================================================================
; ������������� �����
Func _FO_CorrectMask($sMask)
	If StringRegExp($sMask, '[\\/:"<>]') Then Return SetError(2, 0, '|')
	If StringInStr($sMask, '**') Then $sMask = StringRegExpReplace($sMask, '\*+', '*')
	If StringRegExp($sMask & '|', '[\s|.]\|') Then $sMask = StringRegExpReplace($sMask & '|', '[\s|.]+\|', '|')
	If StringInStr('|' & $sMask & '|', '|*|') Then Return '*'
	If $sMask = '|' Then Return SetError(2, 0, '|')
	If StringRight($sMask, 1) = '|' Then $sMask = StringTrimRight($sMask, 1)
	If StringLeft($sMask, 1) = '|' Then $sMask = StringTrimLeft($sMask, 1)
	__FO_MaskUnique($sMask)
	Return $sMask
EndFunc   ;==>_FO_CorrectMask

; �������� ������������� ��������� �����
#Obfuscator_Off
Func __FO_MaskUnique(ByRef $sMask)
	Local $t = StringReplace($sMask, '[', Chr(1)), $a = StringSplit($t, '|'), $k = 0, $i
	Assign('/', '', 1)
	For $i = 1 To $a[0]
		If Not IsDeclared($a[$i] & '/') Then
			$k += 1
			$a[$k] = $a[$i]
		EndIf
		Assign($a[$i] & '/', '', 1)
	Next
	If $k <> $a[0] Then
		$sMask = ''
		For $i = 1 To $k
			$sMask &= $a[$i] & '|'
		Next
		$sMask = StringReplace(StringTrimRight($sMask, 1), Chr(1), '[')
	EndIf
EndFunc   ;==>__FO_MaskUnique
#Obfuscator_On

; ����� ������ �� �����, ������� ����� ������� FileFindFirstFile
; Func __FO_FileSearchMask_Old($sPath, $sMask, $iDepth)
	; Local $sFileList = '', $sFile, $s, $aFolder
	; If $iDepth > 0 Then
		; $aFolder = StringSplit(StringTrimRight($sPath&@CRLF&__FO_FolderSearch($sPath, '*', $iDepth), 2), @CRLF, 1)
	; Else
		; Dim $aFolder[2]=[1, $sPath]
	; EndIf
	; For $i = 1 To $aFolder[0]
		; $s = FileFindFirstFile($aFolder[$i] & $sMask)
		; If $s = -1 Then ContinueLoop
		; While 1
			; $sFile = FileFindNextFile($s)
			; If @error Then ExitLoop
			; If Not @extended Then
				; $sFileList &= $aFolder[$i] & $sFile & @CRLF
			; EndIf
		; WEnd
		; FileClose($s)
	; Next
	; Return $sFileList
; EndFunc

; #FUNCTION# ;=================================================================================
; Function Name ...: _FO_FolderSearch (__FO_FolderSearch, __FO_FolderSearchMask)
; AutoIt Version ....: 3.3.2.0+ , versions below this @extended should be replaced by of StringInStr(FileGetAttrib($sPath&'\'&$sFile), "D")
; Description ........: Search folders on a mask in the subdirectories.
; Syntax................: _FO_FolderSearch ( $sPath [, $sMask = '*' [, $fInclude=True [, $iDepth=0 [, $iFull=1 [, $iArray=1 [, $sLocale=0]]]]]] )
; Parameters:
;		$sPath - search path
;		$sMask - mask using the characters "*" and "?" with the separator "|"
;		$fInclude - (True / False) invert the mask, that is excluded from the search of folders
;		$iDepth - (0-125) nesting level (0 - root directory)
;		$iFull - (0,1)
;                  |0 - relative
;                  |1 - full path
;		$iArray - (0,1,2) if the value other than zero, the result is an array (by default ),
;                  |0 - a list of paths separated by @CRLF
;                  |1 - array, where $iArray[0]=number of folders ( by default)
;                  |2 - array, where $iArray[0] contains the first folder
;		$sLocale - case sensitive.
;                  |0 - not case sensitive (only for 'A-z'), by default.
;                  |1 - case sensitive (for any characters)
;                  |<symbols> - not case sensitive, specified range of characters from local languages. For example '�-���'. 'A-z' is not required, they are enabled by default.
; Return values ....: Success - Array or a list of paths separated by @CRLF
;					Failure - empty string, @error:
;                  |0 - no error
;                  |1 - Invalid path
;                  |2 - Invalid mask
;                  |3 - not found
; Author(s) ..........: AZJIO
; Remarks ..........: Use function _CorrectMask if it is required correct mask, which is entered by user
; ============================================================================================
; ��� ������� ...: _FO_FolderSearch (__FO_FolderSearch, __FO_FolderSearchMask)
; ������ AutoIt ..: 3.3.2.0+ , � ������� ���� ��������� ����� @extended �������� �� StringInStr(FileGetAttrib($sPath&'\'&$sFile), "D")
; �������� ........: ����� ����� �� ����� � ������������.
; ���������.......: _FO_FolderSearch ( $sPath [, $sMask = '*' [, $fInclude=True [, $iDepth=0 [, $iFull=1 [, $iArray=1 [, $sLocale=0]]]]]] )
; ���������:
;		$sPath - ���� ������
;		$sMask - ����� � �������������� �������� "*" � "?" � ������������� ����� "|". �� ��������� ��� �����.
;		$fInclude - (True / False) ������������� �����, �� ���� ��������� �� ������ ��������� �����
;		$iDepth - (0-125) ������� ����������� (0 - �������� �������)
;		$iFull - (0,1)
;                  |0 - �������������
;                  |1 - ������ ����
;		$iArray - (0,1,2) ���������� ����� ����������, ������ ��� ������
;                  |0 - ������ � ������������ @CRLF
;                  |1 - ������, � ������� $iArray[0]=���������� ����� (�� ���������)
;                  |2 - ������, � ������� $iArray[0] �������� ������ �����
;		$sLocale - ��������� ������� ���� ��� ������.
;                  |0 - �� ��������� ������� (������ ��� ��������� ����), �� ���������.
;                  |1 - ��������� ������� (��� ����� ��������)
;                  |<�������> - �� ��������� ������� ���������� ��������� �������� ���������� �����, ������� ���������, �������� '�-���'. ��������� ��������� �� ���������, ��� �� ��������� ��������.
; ������������ ��������: ������� - ������ ��� ������ � ������������ @CRLF
;					�������� - ������ ������, @error:
;                  |0 - ��� ������
;                  |1 - �������� ����
;                  |2 - �������� �����
;                  |3 - ������ �� �������
; ����� ..........: AZJIO
; ���������� ..: ����������� ������� _CorrectMask, ���� ����� ����������� �� ���� ����� � ��������� �������� �� ������������
; ============================================================================================
; ������� �������� � ���������� ������� ���������� � ��������� ��������������� ������
Func _FO_FolderSearch($sPath, $sMask = '*', $fInclude = True, $iDepth = 0, $iFull = 1, $iArray = 1, $sLocale = 0)
	Local $vFolderList, $aFolderList, $i, $rgex
	If $sMask = '|' Then Return SetError(2, 0, '')
	If StringRight($sPath, 1) <> '\' Then $sPath &= '\'
	; If Not StringRegExp($sPath, '(?i)^[a-z]:[^/:*?"<>|]*$') Or StringInStr($sPath, '\\') Then Return SetError(1, 0, '')
	If Not FileExists($sPath) Then Return SetError(1, 0, '')

	If $sMask = '*' Or $sMask = '' Then
		__FO_FolderSearch($vFolderList, $sPath, $iDepth)
		$vFolderList = StringTrimRight($vFolderList, 2)
	Else
		__FO_FolderSearchMask($vFolderList, $sPath, $iDepth)
		$vFolderList = StringTrimRight($vFolderList, 2)
		$sMask = StringReplace(StringReplace(StringRegExpReplace($sMask, '[][$^.{}()+]', '\\$0'), '?', '.'), '*', '.*?')

		If Not $sLocale Then
			$rgex = 'i'
		ElseIf Not($sLocale == '1') Then
			$rgex = 'i'
			$sMask = __FO_UserLocale($sMask, $sLocale)
		EndIf

		If $fInclude Then
			$aFolderList = StringRegExp($vFolderList, '(?m' & $rgex & ')^(.+\|(?:' & $sMask & '))(?:\r|\z)', 3)
			$vFolderList = ''
			For $i = 0 To UBound($aFolderList) - 1
				$vFolderList &= $aFolderList[$i] & @CRLF
			Next
		Else
			$vFolderList = StringRegExpReplace($vFolderList & @CRLF, '(?m' & $rgex & ')^.+\|(' & $sMask & ')\r\n', '')
		EndIf
		$vFolderList = StringReplace(StringTrimRight($vFolderList, 2), '|', '')
	EndIf
	If Not $vFolderList Then Return SetError(3, 0, '')

	If $iFull = 0 Then $vFolderList = StringRegExpReplace($vFolderList, '(?m)^(?:.{' & StringLen($sPath) & '})(.*)$', '\1')
	Switch $iArray
		Case 1
			$vFolderList = StringSplit($vFolderList, @CRLF, 1)
			; If @error And $vFolderList[1] = '' Then Dim $vFolderList[1] = [0]
		Case 2
			$vFolderList = StringSplit($vFolderList, @CRLF, 3)
			; If @error And $vFolderList[0]='' Then SetError(3, 0, '')
	EndSwitch
	Return $vFolderList
EndFunc   ;==>_FO_FolderSearch

; ����� ����� �� �����
Func __FO_FolderSearchMask(ByRef $sFolderList, $sPath, $iDepth, $iCurD = 0)
	Local $sFile, $s = FileFindFirstFile($sPath & '*')
	If $s = -1 Then Return
	While 1
		$sFile = FileFindNextFile($s)
		If @error Then ExitLoop
		If @extended Then
			If $iCurD < $iDepth Then
				$sFolderList &= $sPath & '|' & $sFile & @CRLF
				__FO_FolderSearchMask($sFolderList, $sPath & $sFile & '\', $iDepth, $iCurD + 1)
			ElseIf $iCurD = $iDepth Then
				$sFolderList &= $sPath & '|' & $sFile & @CRLF
			EndIf
		EndIf
	WEnd
	FileClose($s)
EndFunc   ;==>__FO_FolderSearchMask

; ����� ���� �����
Func __FO_FolderSearch(ByRef $sFolderList, $sPath, $iDepth, $iCurD = 0)
	Local $sFile, $s = FileFindFirstFile($sPath & '*')
	If $s = -1 Then Return
	While 1
		$sFile = FileFindNextFile($s)
		If @error Then ExitLoop
		If @extended Then
			If $iCurD < $iDepth Then
				$sFolderList &= $sPath & $sFile & @CRLF
				__FO_FolderSearch($sFolderList, $sPath & $sFile & '\', $iDepth, $iCurD + 1)
			ElseIf $iCurD = $iDepth Then
				$sFolderList &= $sPath & $sFile & @CRLF
			EndIf
		EndIf
	WEnd
	FileClose($s)
EndFunc   ;==>__FO_FolderSearch



; #FUNCTION# ;=================================================================================
; Function Name ...: _FO_SearchEmptyFolders (__FO_SearchEmptyFolders1, __FO_SearchEmptyFolders2)
; AutoIt Version ....: 3.3.2.0+ , versions below this @extended should be replaced by of StringInStr(FileGetAttrib($sPath&'\'&$sFile), "D")
; Description ........: Search for empty folders
; Syntax................: _FO_SearchEmptyFolders ( $sPath [, $iType = 0 [, $iArray = 1 [, $iFull = 1]]] )
; Parameters:
;		$sPath - search path
;		$iType - (0,1) Defines, absolutely empty folders or to allow the catalog with empty folders, without adding nested
;                  |0 - Folder can contain empty folders without adding them to the list (default)
;                  |1 - Folders are empty absolutely
;		$iArray - (0,1,2) if the value other than zero, the result is an array (by default),
;                  |0 - a list of paths separated by @CRLF
;                  |1 - array, where $iArray[0]=number of folders (by default)
;                  |2 - array, where $iArray[0] contains the first folder
;		$iFull - (0,1) Full or relative path
;                  |0 - relative path
;                  |1 - full path (by default)
; Return values ....: Success - Array ($iArray[0]=number of folders) or a list of paths separated by @CRLF
;					Failure - empty string, @error:
;                  |0 - no error
;                  |1 - Invalid path
;                  |2 - not found
;                  |3 - root folder empty
; Author(s) ..........: AZJIO
; Remarks ..........: The main purpose of the function to delete empty folders in the future
; ============================================================================================
; ��� ������� ...: _FO_SearchEmptyFolders (__FO_SearchEmptyFolders1, __FO_SearchEmptyFolders2)
; ������ AutoIt ..: 3.3.2.0+ , � ������� ���� ��������� ����� @extended �������� �� StringInStr(FileGetAttrib($sPath & $sFile), "D")
; �������� ........: ����� ������ �����
; ���������.......: _FO_SearchEmptyFolders ( $sPath [, $iType = 0 [, $iArray = 1 [, $iFull = 1]]] )
; ���������:
;		$sPath - ���� ������
;		$iType - (0,1) ����������, ������ ������ ����� ��� ��������� ������� c ������� �������, �� �������� ���������
;                  |0 - ����� ����� ��������� ������ �����, �� �������� ��������� � ������ (�� ���������)
;                  |1 - ����� ����� ������
;		$iArray - (0,1,2) ���������� ����� ����������, ������ ��� ������
;                  |0 - ������ � ������������ @CRLF
;                  |1 - ������, � ������� $array[0]=���������� ����� (�� ���������)
;                  |2 - ������, � ������� $array[0] �������� ������ �����
;		$iFull - (0,1) ������ ��� ������������� ����
;                  |0 - ������������� ����
;                  |1 - ������ ���� (�� ���������)
; ������������ ��������: ������� - ������ ($array[0]=���������� �����) ��� ������ � ������������ @CRLF
;					�������� - ������ ������, @error:
;                  |0 - ��� ������
;                  |1 - �������� ����
;                  |2 - ������ �� �������
; ����� ..........: AZJIO
; ���������� ..: �������� ���� ������� - ����������� �������� ������ ����� �� ���������� ������
; ============================================================================================
; ����� ������ �����
Func _FO_SearchEmptyFolders($sPath, $iType = 0, $iArray = 1, $iFull = 1)
	If StringRight($sPath, 1) <> '\' Then $sPath &= '\'
	If Not FileExists($sPath) Then Return SetError(1, 0, '')
	Local $sFolderList
	If $iType Then
		$sFolderList = __FO_SearchEmptyFolders1($sPath)
	Else
		$sFolderList = __FO_SearchEmptyFolders2($sPath)
	EndIf
	If Not $sFolderList Then Return SetError(2, 0, '')
	; $sFolderList = StringReplace($sFolderList, '\'&@CR, @CR)
	$sFolderList = StringTrimRight($sFolderList, 2)
	; If $sFolderList = $sPath Then Return SetError(3, 0, '') ;                  |3 - �������� ������� ���� ��� �������� ������ �������� ��� $iType = 0
	If Not $iFull Then $sFolderList = StringRegExpReplace($sFolderList, '(?m)^(?:.{' & StringLen($sPath) & '})(.*)$', '\1')
	Switch $iArray
		Case 1
			$sFolderList = StringSplit($sFolderList, @CRLF, 1)
		Case 2
			$sFolderList = StringSplit($sFolderList, @CRLF, 3)
	EndSwitch
	Return $sFolderList
EndFunc   ;==>_FO_SearchEmptyFolders

Func __FO_SearchEmptyFolders1($sPath)
	Local $sFolderList = '', $sFile, $s = FileFindFirstFile($sPath & '*')
	If $s = -1 Then Return $sPath & @CRLF
	While 1
		$sFile = FileFindNextFile($s)
		If @error Then ExitLoop
		If @extended Then
			$sFolderList &= __FO_SearchEmptyFolders1($sPath & $sFile & '\')
		EndIf
	WEnd
	FileClose($s)
	Return $sFolderList
EndFunc   ;==>__FO_SearchEmptyFolders1

Func __FO_SearchEmptyFolders2($sPath)
	Local $iFill = 0, $sFolderList = '', $sFile, $s = FileFindFirstFile($sPath & '*')
	If $s = -1 Then Return $sPath & @CRLF
	While 1
		$sFile = FileFindNextFile($s)
		If @error Then ExitLoop
		If @extended Then
			$sFolderList &= __FO_SearchEmptyFolders2($sPath & $sFile & '\')
			$iFill += @extended
		Else
			$iFill += 1
		EndIf
	WEnd
	FileClose($s)
	If $iFill = 0 Then
		$s = StringRegExpReplace($sPath, '[][$^.{}()+\\]', '\\$0')
		$sFolderList = StringRegExpReplace($sFolderList, '(?mi)^' & $s & '.*?\r\n', '')
		$sFolderList &= $sPath & @CRLF
	EndIf
	Return SetError(0, $iFill, $sFolderList)
EndFunc   ;==>__FO_SearchEmptyFolders2

; #FUNCTION# ;=================================================================================
; Function Name ...: _FO_FileDirReName
; Description ........: Renaming a file or directory.
; Syntax................: _FO_FileDirReName ( $sSource, $sNewName [, $iFlag=0 [, $iDir=-1 [, $DelAttrib=0]]] )
; Parameters:
;		$sSource - Full path of the file or directory
;		$sNewName - New name
;		$iFlag - (0,1) Flag to overwrite existing
;                  |0 - do not overwrite existing file/directory
;                  |1 - overwrite the existing file (if a directory, preliminary removing)
;		$DelAttrib - (0,1) to remove attributes (-RST), if not let you delete the file/directory
;                  |0 - do not remove the attributes
;                  |1 - remove attributes
;		$iDir - Specifies what is $sSource
;                  |-1 - Auto-detect
;                  |0 - file
;                  |1 - directory
; Return values ....: Success - 1
;					Failure - 0, @error:
;                  |0 - no error
;                  |1 - FileMove or DirMove return failure
;                  |2 - $sNewName - empty string
;                  |3 - $sSource - file / directory in the specified path does not exist
;                  |4 - original and the new name are the same
;                  |5 - $sNewName - contains invalid characters
; Author(s) ..........: AZJIO
; Remarks ..........: If a new file / directory with the same name exists, it will be deleted.
; ============================================================================================
; ��� ������� ...: _FO_FileDirReName
; �������� ........: ��������������� ���� ��� �������.
; ���������.......: _FO_FileDirReName ( $sSource, $sNewName [, $iFlag=0 [, $iDir=-1 [, $DelAttrib=0]]] )
; ���������:
;		$sSource - ������ ���� � �������� ��� �����
;		$sNewName - ����� ���
;		$iFlag - (0,1) ���� ���������� ������������
;                  |0 - �� �������������� ������������ ����/�������
;                  |1 - �������������� ������������ ���� (���� �������, �� ��������������� ��� ��������)
;		$DelAttrib - (0,1) ����� �������� (-RST) ������������� ������� ����/�������
;                  |0 - �� ������� ��������
;                  |1 - ������� ��������
;		$iDir - ��������� ��� �������� $sSource
;                  |-1 - ���������������
;                  |0 - ����
;                  |1 - �������
; ������������ ��������: ������� - 1
;					�������� - 0, @error:
;                  |0 - ��� ������
;                  |1 - FileMove ��� DirMove ���������� �������
;                  |2 - $sNewName - ������ ������
;                  |3 - $sSource - ����/������� �� ���������� ���� �� ����������
;                  |4 - �������� � ����� ��� ���������
;                  |5 - $sNewName - �������� ������������ �������
; ����� ..........: AZJIO
; ���������� ..: ���� ����� ����/������� � ����� �� ������ ����������, �� ����� �����.
; ============================================================================================
Func _FO_FileDirReName($sSource, $sNewName, $iFlag = 0, $DelAttrib = 0, $iDir = -1)
	Local $i, $n, $sName, $sPath, $sTmpPath
	If Not $sNewName Then Return SetError(2, 0, 0)
	If StringRegExp($sNewName, '[\\/:*?"<>|]') Then Return SetError(5, 0, 0) ; (???) ����������� ������ FileMove/DirMove, �� ���� ��������� �� FileMove ����� ������������ �����
	If Not FileExists($sSource) Then Return SetError(3, 0, 0)
	$n = StringInStr($sSource, '\', 0, -1)
	$sPath = StringLeft($sSource, $n)
	$sName = StringTrimLeft($sSource, $n)
	If $iDir = -1 Then $iDir = StringInStr(FileGetAttrib($sSource), 'D')
	$n = 0
	If $sName = $sNewName Then
		If $sName == $sNewName Then Return SetError(4, 0, 0)

		$i = 0
		Do
			$i += 1
			$sTmpPath = $sPath & '$#@_' & $i & '.tmp'
		Until Not FileExists($sTmpPath)

		If $iDir Then
			If DirMove($sSource, $sTmpPath) Then $n = DirMove($sTmpPath, $sPath & $sNewName)
		Else
			If FileMove($sSource, $sTmpPath) Then $n = FileMove($sTmpPath, $sPath & $sNewName)
		EndIf
	Else
		If $iDir Then
			If FileExists($sPath & $sNewName) Then
				If $iFlag Then
					If $DelAttrib Then FileSetAttrib($sPath & $sNewName, "-RST", 1)
					If StringInStr(FileGetAttrib($sPath & $sNewName), 'D') Then
						If DirRemove($sPath & $sNewName, 1) Then $n = DirMove($sSource, $sPath & $sNewName)
					Else
						If FileDelete($sPath & $sNewName) Then $n = DirMove($sSource, $sPath & $sNewName)
					EndIf
				EndIf
			Else
				$n = DirMove($sSource, $sPath & $sNewName)
			EndIf
		Else
			If FileExists($sPath & $sNewName) Then
				If $iFlag Then
					If $DelAttrib Then FileSetAttrib($sPath & $sNewName, "-RST", 1)
					If StringInStr(FileGetAttrib($sPath & $sNewName), 'D') Then
						If DirRemove($sPath & $sNewName, 1) Then $n = FileMove($sSource, $sPath & $sNewName)
					Else
						$n = FileMove($sSource, $sPath & $sNewName, $iFlag)
					EndIf
				EndIf
			Else
				$n = FileMove($sSource, $sPath & $sNewName)
			EndIf
		EndIf
	EndIf
	SetError(Not $n, 0, $n)
EndFunc   ;==>_FO_FileDirReName

; #FUNCTION# ;=================================================================================
; Function Name ...: _FO_GetCopyName
; Description ........: Returns the name of a nonexistent copy of the file
; Syntax................: _FO_GetCopyName ( $sPath [, $iMode=0 [, $sText='Copy']] )
; Parameters:
;		$sPath - Full path of the file or directory
;		$iMode - (0,1) Select the index assignment
;                  |0 - Standard, is similar to creating a copy of a file in Win7
;                  |1 - Append index _1, _2, etc.
;		$sText - Text "Copy"
; Return values ....: The path to the file copy
; Author(s) ..........: AZJIO
; Remarks ..........: There is no error, the function returns the primary name or a new name correctly.
; ============================================================================================
; ��� ������� ...: _FO_GetCopyName
; �������� ........: ���������� ��� �������������� ����� �����.
; ���������.......: _FO_GetCopyName ( $sPath [, $iMode=0 [, $sText='Copy']] )
; ���������:
;		$sPath - ������ ���� � �������� ��� �����
;		$iMode - (0,1) ����� �������� ������������ �������
;                  |0 - �����������, ���������� �������� ����� ����� � Win7
;                  |1 - ���������� ������ ����� _1, _2 � �.�.
;		$sText - ����� "�����", ����� ���� ������ ������������ �� �����������
; ������������ ��������: ���� ����� �����
; ����� ..........: AZJIO
; ���������� ..: ������� �� ���������� ������, ��� ��� ���������� ���� ���������� �� ������ (���� ���� �� ����������), ���� ����� ���������� ���.
; ============================================================================================
Func _FO_GetCopyName($sPath, $iMode = 0, $sText = 'Copy') ; �����, ������� ������������ �� �����������
	Local $i, $aPath[3]
	If FileExists($sPath) Then
		$aPath=_FO_PathSplit($sPath)
		; ���� �������� ���������� ������
		$i = 0
		If $iMode Then
			Do
				$i += 1
				$sPath = $aPath[0] & $aPath[1] & '_' & $i & $aPath[2]
			Until Not FileExists($sPath)
		Else
			Do
				$i += 1
				If $i = 1 Then
					$sPath = $aPath[0] & $aPath[1] & ' ' & $sText & $aPath[2]
				Else
					$sPath = $aPath[0] & $aPath[1] & ' ' & $sText & ' (' & $i & ')' & $aPath[2]
				EndIf
			Until Not FileExists($sPath)
		EndIf
	EndIf
	Return $sPath
EndFunc   ;==>_FO_GetCopyName



; #FUNCTION# ;=================================================================================
; Function Name ...: _FO_FileBackup
; Description ........: Creates a backup file.
; Syntax.......: _FO_FileBackup ( $sPathOriginal [, $sPathBackup='' [, $iCountCopies=3 [, $iDiffSize=0 [, $iDiffTime=0]]]] )
; Parameters:
;		$sPathOriginal - The full path to the original file
;		$sPathBackup - Full or relative path to the backup. The default is "" - empty string, i.e. the current directory
;		$iCountCopies - The maximum number of copies from 1 or more. By default, 3 copies.
;		$iDiffSize - Consider changing the size.
;                  |-1 - Forced to make a reservation
;                  |0 - do not take into consideration the size (by default). In this case, indicate the criterion by date
;                  |1 - to execute reservation if files of the original and a copy are various
;		$iDiffTime - The time interval in seconds between changes in the original and last copies of the file. The default is 0 - do not check.
; Return values: Success - 1, Specifies that the backup performed
;					Failure - 0, @error:
;                  |0 - There is no error, but the backup may fail, with the lack of criteria for reservation
;                  |1 - failed to make a reservation, failure FileCopy or FileMove
;                  |2 - number of copies less than 1
;                  |3 - missing file for backup
; Author(s) ..........: AZJIO
; Remarks ..: The function creates a backup, and the oldest copy is removed. When disabled the criteria (the default) only one copy is created and is not updated in the future.
; ============================================================================================
; ��� ������� ...: _FO_FileBackup
; �������� ........: ������ ��������� ����� �����.
; ���������.......: _FO_FileBackup ( $sPathOriginal [, $sPathBackup='' [, $iCountCopies=3 [, $iDiffSize=0 [, $iDiffTime=0]]]] )
; ���������:
;		$sPathOriginal - ������ ���� � ������������� �����
;		$sPathBackup - ������ ��� ������������� ���� � �������� ��������������. �� ��������� "" - ������ ������, �.� ������� �����
;		$iCountCopies - ������������ ���������� �����, �� 1 � �����. �� ��������� 3 �����.
;		$iDiffSize - (-1, 0, 1) ��������� ��������� �������. ���� 1, �� ����� �� �������� ���� �������� �� ��������� � �������
;                  |-1 - ������������� ������� ��������������
;                  |0 - �� ��������� ������ (�� ���������). � ���� ������ ������� �������� �� ����
;                  |1 - �������������� ����������� ��� �������� �������� ������ ��������� � ��������� ��������� �����
;		$iDiffTime - �������� ������� � �������� ����� ����������� ��������� � ��������� ����� �����. �� ��������� 0 - �� ���������.
; ������������ ��������: ������� - 1, ��������� ��� �������������� ���������
;					�������� - 0, @error:
;                  |0 - ��� ������, �� �������������� ����� �� ����������, ��� ��������� ��������� ��������������
;                  |1 - �� ������� ������� ��������������, ������� FileMove ��� FileCopy
;                  |2 - ���������� ����� ����� 1
;                  |3 - ����������� ���� ��� ��������������
; ����� ..........: AZJIO
; ���������� ..: ������� ������ ��������� ��������� �����, ��� ���� ����� ������ ����� ���������. ��� ����������� ��������� (�� ���������) �������� ������ ���� ����� � �� ����������� � ����������.
; ============================================================================================
Func _FO_FileBackup($sPathOriginal, $sPathBackup = '', $iCountCopies = 3, $iDiffSize = 0, $iDiffTime = 0)
	Local $aPath, $aTB, $aTO, $i, $iDateCalc, $Success
	If $iCountCopies < 1 Then Return SetError(2, 0, 0)
	If Not FileExists($sPathOriginal) Then Return SetError(3, 0, 0)
	$aPath = _FO_PathSplit($sPathOriginal)
	If Not $sPathBackup Then
		$sPathBackup = $aPath[0] ; ���� ������ ������
	ElseIf Not (StringRegExp($sPathBackup, '(?i)^[a-z]:[^/:*?"<>|]*$') Or StringInStr($sPathBackup, '\\')) Then ; ���� �� ������ ���� ��� �� UNC
		If StringRegExp($sPathBackup, '[/:*?"<>|]') Then
			$sPathBackup = $aPath[0]
		Else
			$sPathBackup = StringReplace($aPath[0] & $sPathBackup & '\', '\\', '\') ; �� ������������� ����
		EndIf
	EndIf
	Switch $iDiffSize
		Case -1
			$iDiffSize = 1
		Case 0
			$iDiffSize = 0
		Case Else
			$iDiffSize = (FileGetSize($sPathOriginal) <> FileGetSize($sPathBackup & $aPath[1] & '_1' & $aPath[2]))
	EndSwitch
	If $iDiffTime Then
		$aTO = FileGetTime($sPathOriginal)
		$aTB = FileGetTime($sPathBackup & $aPath[1] & '_1' & $aPath[2])
		If Not @error Then
			$iDateCalc = _DateDiff('s', $aTB[0] & '/' & $aTB[1] & '/' & $aTB[2] & ' ' & $aTB[3] & ':' & $aTB[4] & ':' & $aTB[5], $aTO[0] & '/' & $aTO[1] & '/' & $aTO[2] & ' ' & $aTO[3] & ':' & $aTO[4] & ':' & $aTO[5])
			$iDiffTime = ($iDateCalc > $iDiffTime)
		EndIf
	EndIf
	$sPathBackup &= $aPath[1]
	If Not FileExists($sPathBackup & '_1' & $aPath[2]) Or $iDiffSize Or $iDiffTime Then
		$Success = 1
		For $i = $iCountCopies To 2 Step -1
			If $Success And FileExists($sPathBackup & '_' & $i - 1 & $aPath[2]) Then $Success = FileMove($sPathBackup & '_' & $i - 1 & $aPath[2], $sPathBackup & '_' & $i & $aPath[2], 9)
		Next
		If $Success Then $Success = FileCopy($sPathOriginal, $sPathBackup & '_1' & $aPath[2], 9)
		Return SetError(Not $Success, 0, $Success)
	EndIf
	Return SetError(0, 0, 0)
EndFunc   ;==>_FO_FileBackup

; #FUNCTION# ;=================================================================================
; Function Name ...: _FO_PathSplit
; Description ........: Divides the path into 3 parts : path, file, and extension.
; Syntax.......: _FO_PathSplit ( $sPath )
; Parameters:
;		$sPath - Path
; Return values: Success - array in the following format
;		$Array[0] = path
;		$Array[1] = name of the file / directory
;		$Array[2] = extension
; Author(s) ..........: AZJIO
; Remarks ..: Function has no errors. If you do not have any portion of a path, the array contains an empty cell for this item
; ============================================================================================
; ��� ������� ...: _FO_PathSplit
; �������� ........: ����� ���� �� 3 �����: ����, ����, ����������.
; ���������.......: _FO_PathSplit ( $sPath )
; ���������:
;		$sPath - ����
; ������������ ��������: ������� - ������ �� 3-x ��������� ���������� �������
;		$Array[0] = ����
;		$Array[1] = ��� ����� / ��������
;		$Array[2] = ����������
; ����� ..........: AZJIO
; ���������� ..: ������� �� ����� ������. ���� ����������� ����� ���� ������� ����, �� ������ �������� ������ ������ ��� ����� ��������
; ============================================================================================
Func _FO_PathSplit($sPath)
	Local $i, $aPath[3] ; ( Dir | Name | Ext )
	$i = StringInStr($sPath, '\', 0, -1)
	$aPath[1] = StringTrimLeft($sPath, $i)
	$aPath[0] = StringLeft($sPath, $i) ; Dir
	If StringInStr($aPath[1], '.') Then
		$i = StringInStr($aPath[1], '.', 0, -1) - 1
		$aPath[2] = StringTrimLeft($aPath[1], $i) ; Ext
		$aPath[1] = StringLeft($aPath[1], $i) ; Name
	Else
		$aPath[2] = ''
	EndIf
	Return $aPath
EndFunc   ;==>_FO_PathSplit