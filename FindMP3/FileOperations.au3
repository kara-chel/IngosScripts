#include-once
#include <Date.au3> ; _FO_FileBackup

; FileOperations.au3
; версия 1.8.5 от 2014.05.26

; http://www.autoitscript.com/forum/topic/133224-filesearch-foldersearch/
; http://pastebin.com/AbzMyg1x
; http://azjio.ucoz.ru/publ/skripty_autoit3/funkcii/filesearch/11-1-0-33

; =======================================
; Title .........: FileOperations
; AutoIt Version : 3.3.2.0 - 3.3.8.1
; Language ......: English + Русский
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
; _FO_IsDir
; _FO_ShortFileSize
; _FO_IsEmptyFolder
; =======================================

; Внутренние функции
; #INTERNAL_USE_ONLY#====================
; __FO_FileSearchType
; __FO_FileSearchMask
; __FO_FileSearchAll
; __FO_GetListMask
; __FO_MaskUnique (#Obfuscator_Off и #Obfuscator_On)
; __FO_FolderSearch
; __FO_FolderSearchMask
; __FO_SearchEmptyFolders1
; __FO_SearchEmptyFolders2
; __FO_UserLocale
; __ChExcludeFolders
; =======================================

; #FUNCTION# ;=================================================================================
; Function Name ...: _FO_FileSearch (__FO_FileSearchType, __FO_FileSearchMask, __FO_FileSearchAll)
; AutoIt Version ....: 3.3.2.0+ , versions below this @extended should be replaced by of StringInStr(FileGetAttrib($sPath&'\'&$sFile), "D")
; Description ........: Search files by mask in subdirectories.
; Syntax................: _FO_FileSearch ( $sPath [, $sMask = '*' [, $fInclude=True [, $iDepth=125 [, $iFull=1 [, $iArray=1 [, $iTypeMask=1 [, $sLocale=0[, $vExcludeFolders = ''[, $iExcludeDepth = -1]]]]]]]]] )
; Parameters:
;		$sPath - Search path
;		$sMask - Allowed two options for the mask: using symbols "*" and "?" with the separator "|", or a list of extensions with the separator "|"
;		$fInclude - (True / False) Invert the mask, that is excluded from the search for these types of files
;		$iDepth - (0-125) Nesting level (0 - root directory)
;		$iFull - (0,1,2,3)
;                  |0 - Relative
;                  |1 - Full path
;                  |2 - File names with extension
;                  |3 - File names without extension
;		$iArray - if the value other than zero, the result is an array (by default ),
;                  |0 - A list of paths separated by @CRLF
;                  |1 - Array, where $iArray[0]=number of files ( by default)
;                  |2 - Array, where $iArray[0] contains the first file
;		$iTypeMask - (0,1,2) defines the format mask
;                  |0 - Auto detect
;                  |1 - Forced mask, for example *.is?|s*.cp* (it is possible to specify a file name with no characters * or ? and no extension will be found)
;                  |2 - Forced mask, for example tmp|bak|gid (that is, only files with the specified extension)
;		$sLocale - Case sensitive.
;                  |-1 - Not case sensitive (only for 'A-z').
;                  |0 - Not case sensitive, by default. (for any characters)
;                  |1 - Case sensitive (for any characters)
;                  |<symbols> - not case sensitive, specified range of characters from local languages. For example 'А-яЁё'. 'A-z' is not required, they are enabled by default.
;		$vExcludeFolders - Excludes folders from search. List the folder names via the "|", for example, "Name1|Name2|Name3|".
;		$iExcludeDepth - Nesting level for the parameter $vExcludeFolders. -1 by default, which means disabled.
; Return values ....: Success - Array or a list of paths separated by @CRLF
;					Failure - Empty string, @error:
;                  |0 - No error
;                  |1 - Invalid path
;                  |2 - Invalid mask
;                  |3 - Not found
; Author(s) ..........: AZJIO
; Remarks ..........: Use function _CorrectMask if it is required correct mask, which is entered by user
; ============================================================================================
; Имя функции ...: _FO_FileSearch (__FO_FileSearchType, __FO_FileSearchMask, __FO_FileSearchAll)
; Версия AutoIt ..: 3.3.2.0+ , в версиях ниже указанной нужно @extended заменить на StringInStr(FileGetAttrib($sPath&'\'&$sFile), "D")
; Описание ........: Поиск файлов по маске в подкаталогах.
; Синтаксис.......: _FO_FileSearch ( $sPath [, $sMask = '*' [, $fInclude=True [, $iDepth=125 [, $iFull=1 [, $iArray=1 [, $iTypeMask=1 [, $sLocale=0[, $vExcludeFolders = ''[, $iExcludeDepth = -1]]]]]]]]] )
; Параметры:
;		$sPath - Путь поиска
;		$sMask - Допустимы два варианта маски: с использованием символов "*" и "?" с перечислением через "|", либо перечисление расширений через "|"
;		$fInclude - (True / False) Инвертировать маску, то есть исключить из поиска указанные типы файлов
;		$iDepth - (0-125) Уровень вложенности (0 - корневой каталог)
;		$iFull - (0,1,2,3)
;                  |0 - Относительный
;                  |1 - Полный путь
;                  |2 - Имена файлов с расширением
;                  |3 - Имена файлов без расширения
;		$iArray - (0,1,2) Определяет вывод результата, массив или список
;                  |0 - Список с разделителем @CRLF
;                  |1 - Массив, в котором $iArray[0]=количество файлов (по умолчанию)
;                  |2 - Массив, в котором $iArray[0] содержит первый файл
;		$iTypeMask - (0,1,2) Определяет формат записи маски
;                  |0 - Автоопределение
;                  |1 - Принудительно маска вида *.is?|s*.cp* (то есть можно указать имя файла без символов * или ? и без расширения и будет найдено)
;                  |2 - Принудительно маска вида tmp|bak|gid (по расширению, то есть только имена файлов с указанным расширением)
;		$sLocale - Учитывать регистр букв при поиске.
;                  |-1 - Не учитывать регистр только для латинских букв
;                  |0 - Не учитывать регистр, по умолчанию (для всех символов).
;                  |1 - Учитывать регистр (для всех символов).
;                  |<символы> - не учитывать регистр указанного диапазона символов локального языка, включая латинские, например 'А-яЁё'. Латинские указывать не требуется, они по умолчанию включены.
;		$vExcludeFolders - Исключает папки из поиска. Перечислять имена папок через "|", например "имя1|имя2|имя3".
;		$iExcludeDepth - Уровень вложенности, до которого работает исключение, указанное в параметре $vExcludeFolders. По умолчанию -1, что означает отключено.
; Возвращаемое значение: Успешно - Массив или список с разделителем @CRLF
;					Неудачно - пустая строка, @error:
;                  |0 - Нет ошибок
;                  |1 - Неверный путь
;                  |2 - Неверная маска
;                  |3 - Ничего не найдено
; Автор ..........: AZJIO
; Примечания ..: Используйте функцию _CorrectMask, если маска считывается из поля ввода и требуется проверка на корректность
; ============================================================================================
; функция проверки и подготовки входных параметров и обработка результирующего списка
Func _FO_FileSearch($sPath, $sMask = '*', $fInclude = True, $iDepth = 125, $iFull = 1, $iArray = 1, $iTypeMask = 1, $sLocale = 0, $vExcludeFolders = '', $iExcludeDepth = -1)
	Local $vFileList
	If $sMask = '|' Then Return SetError(2, 0, '')
	; If Not StringRegExp($sPath, '(?i)^[a-z]:[^/:*?"<>|]*$') Or StringInStr($sPath, '\\') Then Return SetError(1, 0, '')
	If Not FileExists($sPath) Then Return SetError(1, 0, '')
	If StringRight($sPath, 1) <> '\' Then $sPath &= '\'

	If $vExcludeFolders Then
		$vExcludeFolders = StringSplit($vExcludeFolders, '|')
	Else
		Dim $vExcludeFolders[1] = [0]
	EndIf

	If $sMask = '*' Or $sMask = '' Then
		__FO_FileSearchAll($vFileList, $sPath, $iDepth, $vExcludeFolders, $iExcludeDepth)
	Else
		Switch $iTypeMask
			Case 0
				If StringInStr($sMask, '*') Or StringInStr($sMask, '?') Or StringInStr($sMask, '.') Then
					__FO_GetListMask($sPath, $sMask, $fInclude, $iDepth, $vFileList, $sLocale, $vExcludeFolders, $iExcludeDepth)
				Else
					__FO_FileSearchType($vFileList, $sPath, '|' & $sMask & '|', $fInclude, $iDepth, $vExcludeFolders, $iExcludeDepth)
				EndIf
			Case 1
				__FO_GetListMask($sPath, $sMask, $fInclude, $iDepth, $vFileList, $sLocale, $vExcludeFolders, $iExcludeDepth)
			Case Else
				If StringInStr($sMask, '*') Or StringInStr($sMask, '?') Or StringInStr($sMask, '.') Then Return SetError(2, 0, '')
				__FO_FileSearchType($vFileList, $sPath, '|' & $sMask & '|', $fInclude, $iDepth, $vExcludeFolders, $iExcludeDepth)
		EndSwitch
	EndIf

	If Not $vFileList Then Return SetError(3, 0, '')
	Switch $iFull
		Case 0
			$vFileList = StringRegExpReplace($vFileList, '(?m)^[^\v]{' & StringLen($sPath) & '}', '')
		Case 2
			$vFileList = StringRegExpReplace($vFileList, '(?m)^.*\\', '')
		Case 3
			$vFileList = StringRegExpReplace($vFileList, '(?m)^[^\v]+\\', '')
			$vFileList = StringRegExpReplace($vFileList, '(?m)\.[^./:*?"<>|\\\v]+\r?$', @CR)
	EndSwitch
	$vFileList = StringTrimRight($vFileList, 2)
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

; Получение списка и обработка регулярным выражением
Func __FO_GetListMask($sPath, $sMask, $fInclude, $iDepth, ByRef $sFileList, $sLocale, ByRef $aExcludeFolders, ByRef $iExcludeDepth)
	Local $aFileList, $rgex
	__FO_FileSearchMask($sFileList, $sPath, $iDepth, $aExcludeFolders, $iExcludeDepth)
	$sFileList = StringTrimRight($sFileList, 2)
	$sMask = StringReplace(StringReplace(StringRegExpReplace($sMask, '[][$^.{}()+]', '\\$0'), '?', '.'), '*', '.*?')

	Switch $sLocale
		Case -1 ;  не учитывать регистр только для латинских букв
			$rgex = 'i'
		Case 1 ; как есть '', учитывать регистр
		Case 0 ;  не учитывать регистр, по умолчанию.
			$sLocale = '\x{80}-\x{ffff}'
			ContinueCase
		Case Else ; Иначе указан диапазон или набор символов
			$rgex = 'i'
			$sMask = __FO_UserLocale($sMask, $sLocale)
	EndSwitch

	If $fInclude Then
		$aFileList = StringRegExp($sFileList, '(?m' & $rgex & ')^([^|]+\|(?:' & $sMask & '))(?:\r|\z)', 3)
		$sFileList = ''
		For $i = 0 To UBound($aFileList) - 1
			$sFileList &= $aFileList[$i] & @CRLF
		Next
	Else
		$sFileList = StringRegExpReplace($sFileList & @CRLF, '(?m' & $rgex & ')^[^|]+\|(' & $sMask & ')\r\n', '')
	EndIf
	$sFileList = StringReplace($sFileList, '|', '')
EndFunc   ;==>__FO_GetListMask

Func __FO_UserLocale($sMask, $sLocale)
	Local $s, $tmp
	$sLocale = StringRegExpReplace($sMask, '[^' & $sLocale & ']', '')
	$tmp = StringLen($sLocale)
	For $i = 1 To $tmp
		$s = StringMid($sLocale, $i, 1)
		If $s Then
			If StringInStr($sLocale, $s, 0, 2, $i) Then
				$sLocale = $s & StringReplace($sLocale, $s, '')
			EndIf
		Else
			ExitLoop
		EndIf
	Next
	If $sLocale Then
		Local $Upper, $Lower
		$tmp = StringSplit($sLocale, '')
		For $i = 1 To $tmp[0]
			$Upper = StringUpper($tmp[$i])
			$Lower = StringLower($tmp[$i])
			If Not ($Upper == $Lower) Then $sMask = StringReplace($sMask, $tmp[$i], '[' & $Upper & $Lower & ']')
		Next
	EndIf
	Return $sMask
EndFunc   ;==>__FO_UserLocale

; поиск указанных типов файлов
Func __FO_FileSearchType(ByRef $sFileList, $sPath, $sMask, ByRef $fInclude, ByRef $iDepth, ByRef $aExcludeFolders, ByRef $iExcludeDepth, $iCurD = 0)
	Local $iPos, $sFile, $s = FileFindFirstFile($sPath & '*')
	If $s = -1 Then Return
	While 1
		$sFile = FileFindNextFile($s)
		If @error Then ExitLoop
		If @extended Then
			If $iCurD >= $iDepth Or ($iCurD <= $iExcludeDepth And __ChExcludeFolders($sFile, $aExcludeFolders)) Then ContinueLoop
			__FO_FileSearchType($sFileList, $sPath & $sFile & '\', $sMask, $fInclude, $iDepth, $aExcludeFolders, $iExcludeDepth, $iCurD + 1)
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

; поиск файлов по маске
Func __FO_FileSearchMask(ByRef $sFileList, $sPath, ByRef $iDepth, ByRef $aExcludeFolders, ByRef $iExcludeDepth, $iCurD = 0)
	Local $sFile, $s = FileFindFirstFile($sPath & '*')
	If $s = -1 Then Return
	While 1
		$sFile = FileFindNextFile($s)
		If @error Then ExitLoop
		If @extended Then
			If $iCurD >= $iDepth Or ($iCurD <= $iExcludeDepth And __ChExcludeFolders($sFile, $aExcludeFolders)) Then ContinueLoop
			__FO_FileSearchMask($sFileList, $sPath & $sFile & '\', $iDepth, $aExcludeFolders, $iExcludeDepth, $iCurD + 1)
		Else
			$sFileList &= $sPath & '|' & $sFile & @CRLF
		EndIf
	WEnd
	FileClose($s)
EndFunc   ;==>__FO_FileSearchMask

; поиск всех файлов
Func __FO_FileSearchAll(ByRef $sFileList, $sPath, ByRef $iDepth, ByRef $aExcludeFolders, ByRef $iExcludeDepth, $iCurD = 0)
	Local $sFile, $s = FileFindFirstFile($sPath & '*')
	If $s = -1 Then Return
	While 1
		$sFile = FileFindNextFile($s)
		If @error Then ExitLoop
		If @extended Then
			If $iCurD >= $iDepth Or ($iCurD <= $iExcludeDepth And __ChExcludeFolders($sFile, $aExcludeFolders)) Then ContinueLoop
			__FO_FileSearchAll($sFileList, $sPath & $sFile & '\', $iDepth, $aExcludeFolders, $iExcludeDepth, $iCurD + 1)
		Else
			$sFileList &= $sPath & $sFile & @CRLF
		EndIf
	WEnd
	FileClose($s)
EndFunc   ;==>__FO_FileSearchAll

Func __ChExcludeFolders(ByRef $sFile, ByRef $aExcludeFolders)
	For $i = 1 To $aExcludeFolders[0]
		If $sFile = $aExcludeFolders[$i] Then Return True
	Next
	Return False
EndFunc   ;==>__ChExcludeFolders

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
; Имя функции ...: _FO_CorrectMask (__FO_MaskUnique)
; Версия AutoIt ..: 3.3.0.0+
; Описание ........: Корректировка маски
; Синтаксис.......: _CorrectMask ( $sMask )
; Параметры.....: $sMask - кроме символов допустимых в именах допускаются символы подстановки "*" и "?" и разделитель "|"
; Возвращаемое значение:
;					|Успешно -  Возвращает строку корректной маски
;					|Неудачно - Возвращает символ "|" и @error=2
; Автор ..........: AZJIO
; Примечания ..: Функция исправляет возможные ошибки ввода пользователем:
; удаляет пробелы и точки на конце каждого элемента маски, удаляет повторы звёздочки и разделителя.
; ============================================================================================
; корректировка маски
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

; удаление повторяющихся элементов маски
#Obfuscator_Off
Func __FO_MaskUnique(ByRef $sMask)
	Local $t = StringReplace($sMask, '[', Chr(1)), $a = StringSplit($t, '|'), $k = 0
	Assign('/', '', 1)
	For $i = 1 To $a[0]
		If Not IsDeclared($a[$i] & '/') Then
			$k += 1
			$a[$k] = $a[$i]
			Assign($a[$i] & '/', '', 1)
		EndIf
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

; #FUNCTION# ;=================================================================================
; Function Name ...: _FO_FolderSearch (__FO_FolderSearch, __FO_FolderSearchMask)
; AutoIt Version ....: 3.3.2.0+ , versions below this @extended should be replaced by of StringInStr(FileGetAttrib($sPath&'\'&$sFile), "D")
; Description ........: Search folders on a mask in the subdirectories.
; Syntax................: _FO_FolderSearch ( $sPath [, $sMask = '*' [, $fInclude=True [, $iDepth=0 [, $iFull=1 [, $iArray=1 [, $sLocale=0]]]]]] )
; Parameters:
;		$sPath - Search path
;		$sMask - Mask using the characters "*" and "?" with the separator "|"
;		$fInclude - (True / False) invErt the mask, that is excluded from the search of folders
;		$iDepth - (0-125) Nesting level (0 - root directory)
;		$iFull - (0,1)
;                  |0 - Relative
;                  |1 - Full path
;		$iArray - (0,1,2) If the value other than zero, the result is an array (by default ),
;                  |0 - A list of paths separated by @CRLF
;                  |1 - Array, where $iArray[0]=number of folders ( by default)
;                  |2 - Array, where $iArray[0] contains the first folder
;		$sLocale - Case sensitive.
;                  |-1 - Not case sensitive (only for 'A-z').
;                  |0 - Not case sensitive, by default. (for any characters)
;                  |1 - Case sensitive (for any characters)
;                  |<symbols> - Not case sensitive, specified range of characters from local languages. For example 'А-яЁё'. 'A-z' is not required, they are enabled by default.
; Return values ....: Success - Array or a list of paths separated by @CRLF
;					Failure - Empty string, @error:
;                  |0 - No error
;                  |1 - Invalid path
;                  |2 - Invalid mask
;                  |3 - Not found
; Author(s) ..........: AZJIO
; Remarks ..........: Use function _CorrectMask if it is required correct mask, which is entered by user
; ============================================================================================
; Имя функции ...: _FO_FolderSearch (__FO_FolderSearch, __FO_FolderSearchMask)
; Версия AutoIt ..: 3.3.2.0+ , в версиях ниже указанной нужно @extended заменить на StringInStr(FileGetAttrib($sPath&'\'&$sFile), "D")
; Описание ........: Поиск папок по маске в подкаталогах.
; Синтаксис.......: _FO_FolderSearch ( $sPath [, $sMask = '*' [, $fInclude=True [, $iDepth=0 [, $iFull=1 [, $iArray=1 [, $sLocale=0]]]]]] )
; Параметры:
;		$sPath - Путь поиска
;		$sMask - Маска с использованием символов "*" и "?" с перечислением через "|". По умолчанию все папки.
;		$fInclude - (True / False) Инвертировать маску, то есть исключить из поиска указанные папки
;		$iDepth - (0-125) Уровень вложенности (0 - корневой каталог)
;		$iFull - (0,1)
;                  |0 - Относительный
;                  |1 - Полный путь
;		$iArray - (0,1,2) Определяет вывод результата, массив или список
;                  |0 - Список с разделителем @CRLF
;                  |1 - Массив, в котором $iArray[0]=количество папок (по умолчанию)
;                  |2 - Массив, в котором $iArray[0] содержит первую папку
;		$sLocale - Учитывать регистр букв при поиске.
;                  |-1 - Не учитывать регистр только для латинских букв
;                  |0 - Не учитывать регистр, по умолчанию (для всех символов).
;                  |1 - Учитывать регистр (для всех символов).
;                  |<символы> - Не учитывать регистр указанного диапазона символов локального языка, включая латинские, например 'А-яЁё'. Латинские указывать не требуется, они по умолчанию включены.
; Возвращаемое значение: Успешно - Массив или список с разделителем @CRLF
;					Неудачно - Пустая строка, @error:
;                  |0 - Нет ошибок
;                  |1 - Неверный путь
;                  |2 - Неверная маска
;                  |3 - Ничего не найдено
; Автор ..........: AZJIO
; Примечания ..: Используйте функцию _CorrectMask, если маска считывается из поля ввода и требуется проверка на корректность
; ============================================================================================
; функция проверки и подготовки входных параметров и обработка результирующего списка
Func _FO_FolderSearch($sPath, $sMask = '*', $fInclude = True, $iDepth = 0, $iFull = 1, $iArray = 1, $sLocale = 0)
	Local $vFolderList, $aFolderList, $rgex
	If $sMask = '|' Then Return SetError(2, 0, '')
	; If Not StringRegExp($sPath, '(?i)^[a-z]:[^/:*?"<>|]*$') Or StringInStr($sPath, '\\') Then Return SetError(1, 0, '')
	If Not FileExists($sPath) Then Return SetError(1, 0, '')
	If StringRight($sPath, 1) <> '\' Then $sPath &= '\'

	If $sMask = '*' Or $sMask = '' Then
		__FO_FolderSearch($vFolderList, $sPath, $iDepth)
		$vFolderList = StringTrimRight($vFolderList, 2)
	Else
		__FO_FolderSearchMask($vFolderList, $sPath, $iDepth)
		$vFolderList = StringTrimRight($vFolderList, 2)
		$sMask = StringReplace(StringReplace(StringRegExpReplace($sMask, '[][$^.{}()+]', '\\$0'), '?', '.'), '*', '.*?')

		Switch $sLocale
			Case -1 ;  не учитывать регистр только для латинских букв
				$rgex = 'i'
			Case 1 ; как есть '', учитывать регистр
			Case 0 ;  не учитывать регистр, по умолчанию.
				$sLocale = '\x{80}-\x{ffff}'
				ContinueCase
			Case Else ; Иначе указан диапазон или набор символов
				$rgex = 'i'
				$sMask = __FO_UserLocale($sMask, $sLocale)
		EndSwitch

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

	If $iFull = 0 Then $vFolderList = StringRegExpReplace($vFolderList, '(?m)^[^\v]{' & StringLen($sPath) & '}', '')

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

; поиск папок по маске
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

; поиск всех папок
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
; Author(s) ..........: AZJIO
; Remarks ..........: The main purpose of the function to delete empty folders in the future
; ============================================================================================
; Имя функции ...: _FO_SearchEmptyFolders (__FO_SearchEmptyFolders1, __FO_SearchEmptyFolders2)
; Версия AutoIt ..: 3.3.2.0+ , в версиях ниже указанной нужно @extended заменить на StringInStr(FileGetAttrib($sPath & $sFile), "D")
; Описание ........: Поиск пустых папок
; Синтаксис.......: _FO_SearchEmptyFolders ( $sPath [, $iType = 0 [, $iArray = 1 [, $iFull = 1]]] )
; Параметры:
;		$sPath - путь поиска
;		$iType - (0,1) определяет, совсем пустые папки или разрешить каталог c пустыми папками, не добавляя вложенных
;                  |0 - папка может содержать пустые папки, не добавляя последние к списку (по умолчанию)
;                  |1 - папки пусты совсем
;		$iArray - (0,1,2) определяет вывод результата, массив или список
;                  |0 - список с разделителем @CRLF
;                  |1 - массив, в котором $array[0]=количество папок (по умолчанию)
;                  |2 - массив, в котором $array[0] содержит первую папку
;		$iFull - (0,1) Полный или относительный путь
;                  |0 - относительный путь
;                  |1 - полный путь (по умолчанию)
; Возвращаемое значение: Успешно - Массив ($array[0]=количество папок) или список с разделителем @CRLF
;					Неудачно - пустая строка, @error:
;                  |0 - нет ошибок
;                  |1 - неверный путь
;                  |2 - ничего не найдено
; Автор ..........: AZJIO
; Примечания ..: Основная цель функции - последующее удаление пустых папок по найденному списку
; ============================================================================================
; поиск пустых папок
Func _FO_SearchEmptyFolders($sPath, $iType = 0, $iArray = 1, $iFull = 1)
	If Not _FO_IsDir($sPath) Then Return SetError(1, 0, '')
	If StringRight($sPath, 1) <> '\' Then $sPath &= '\'
	Local $sFolderList
	If $iType Then
		$sFolderList = __FO_SearchEmptyFolders1($sPath)
	Else
		$sFolderList = __FO_SearchEmptyFolders2($sPath)
	EndIf
	If Not $sFolderList Then Return SetError(2, 0, '')
	; $sFolderList = StringReplace($sFolderList, '\'&@CR, @CR)
	$sFolderList = StringTrimRight($sFolderList, 2)
	; If $sFolderList = $sPath Then Return SetError(3, 0, '') ;                  |3 - корневой каталог пуст или содержит пустые каталоги при $iType = 0
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
; Syntax................: _FO_FileDirReName ( $sSource, $sNewName [, $iFlag=0 [, $DelAttrib=0 [, $iDir=-1]]] )
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
; Имя функции ...: _FO_FileDirReName
; Описание ........: Переименовывает файл или каталог.
; Синтаксис.......: _FO_FileDirReName ( $sSource, $sNewName [, $iFlag=0 [, $DelAttrib=0 [, $iDir=-1]]] )
; Параметры:
;		$sSource - Полный путь к каталогу или файлу
;		$sNewName - Новое имя
;		$iFlag - (0,1) Флаг перезаписи существующих
;                  |0 - не перезаписывать существующий файл/каталог
;                  |1 - перезаписывать существующий файл (если каталог, то предварительное его удаление)
;		$DelAttrib - (0,1) Снять атрибуты (-RST) непозволяющие удалить файл/каталог
;                  |0 - не снимать атрибуты
;                  |1 - снимать атрибуты
;		$iDir - Указывает чем является $sSource
;                  |-1 - автоопределение
;                  |0 - файл
;                  |1 - каталог
; Возвращаемое значение: Успешно - 1
;					Неудачно - 0, @error:
;                  |0 - нет ошибок
;                  |1 - FileMove или DirMove возвратили неудачу
;                  |2 - $sNewName - пустая строка
;                  |3 - $sSource - файл/каталог по указанному пути не существует
;                  |4 - исходное и новое имя совпадают
;                  |5 - $sNewName - содержит недопустимые символы
; Автор ..........: AZJIO
; Примечания ..: Если новый файл/каталог с таким же именем существует, то будет удалён.
; ============================================================================================
Func _FO_FileDirReName($sSource, $sNewName, $iFlag = 0, $DelAttrib = 0, $iDir = -1)
	Local $i, $n, $sName, $sPath, $sTmpPath
	If Not $sNewName Then Return SetError(2, 0, 0)
	If StringRegExp($sNewName, '[\\/:*?"<>|]') Then Return SetError(5, 0, 0) ; (???) Индикатором служит FileMove/DirMove, но если отключить то FileMove будет поддерживать маски
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
; Имя функции ...: _FO_GetCopyName
; Описание ........: Возвращает имя несуществующей копии файла.
; Синтаксис.......: _FO_GetCopyName ( $sPath [, $iMode=0 [, $sText='Copy']] )
; Параметры:
;		$sPath - Полный путь к каталогу или файлу
;		$iMode - (0,1) Выбор варианта присваивания индекса
;                  |0 - Стандартное, аналогично созданию копии файла в Win7
;                  |1 - Дописывать индекс копии _1, _2 и т.д.
;		$sText - Текст "Копия", может быть изменён взависимости от локализации
; Возвращаемое значение: Путь копии файла
; Автор ..........: AZJIO
; Примечания ..: Функция не возвращает ошибки, так как возвращает либо переданную ей строку (если файл не существует), либо новое корректное имя.
; ============================================================================================
Func _FO_GetCopyName($sPath, $iMode = 0, $sText = 'Copy') ; Копия, префикс взависимости от локализации
	Local $i, $aPath[3]
	If FileExists($sPath) Then
		$aPath = _FO_PathSplit($sPath)
		; цикл проверки одноимённых файлов
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
; Имя функции ...: _FO_FileBackup
; Описание ........: Создаёт резервную копию файла.
; Синтаксис.......: _FO_FileBackup ( $sPathOriginal [, $sPathBackup='' [, $iCountCopies=3 [, $iDiffSize=0 [, $iDiffTime=0]]]] )
; Параметры:
;		$sPathOriginal - Полный путь к оригинальному файлу
;		$sPathBackup - Полный или относительный путь к каталогу резервирования. По умолчанию "" - пустая строка, т.е текущая папка
;		$iCountCopies - Максимальное количество копий, от 1 и более. По умолчанию 3 копии.
;		$iDiffSize - (-1, 0, 1) Учитывать изменение размера. Если 1, то копия не создаётся если оригинал не изменился в размере
;                  |-1 - принудительно сделать резервирование
;                  |0 - не учитывать размер (по умолчанию). В этом случае указать критерий по дате
;                  |1 - резервирование выполняется при различии размеров файлов оригинала и последней созданной копии
;		$iDiffTime - Интервал времени в секундах между изменениями оригинала и последней копии файла. По умолчанию 0 - не проверять.
; Возвращаемое значение: Успешно - 1, указывает что резервирование выполнено
;					Неудачно - 0, @error:
;                  |0 - нет ошибок, но резервирование может не выполнится, при отсутсвии критериев резервирования
;                  |1 - не удалось сделать резервирование, неудача FileMove или FileCopy
;                  |2 - количество копий менее 1
;                  |3 - отсутствует файл для резервирования
; Автор ..........: AZJIO
; Примечания ..: Функция создаёт очередную резервную копию, при этом самая старая копия удаляется. При отключенных критериях (по умолчанию) создаётся только одна копия и не обновляется в дальнейшем.
; ============================================================================================
Func _FO_FileBackup($sPathOriginal, $sPathBackup = '', $iCountCopies = 3, $iDiffSize = 0, $iDiffTime = 0)
	Local $aPath, $aTB, $aTO, $iDateCalc, $Success
	If $iCountCopies < 1 Then Return SetError(2, 0, 0)
	If Not FileExists($sPathOriginal) Then Return SetError(3, 0, 0)
	$aPath = _FO_PathSplit($sPathOriginal)
	If Not $sPathBackup Then
		$sPathBackup = $aPath[0] ; если пустая строка
	ElseIf Not (StringRegExp($sPathBackup, '(?i)^[a-z]:[^/:*?"<>|]*$') Or StringInStr($sPathBackup, '\\')) Then ; если не полный путь или не UNC
		If StringRegExp($sPathBackup, '[/:*?"<>|]') Then
			$sPathBackup = $aPath[0]
		Else
			$sPathBackup = StringReplace($aPath[0] & $sPathBackup & '\', '\\', '\') ; то относительный путь
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
; Имя функции ...: _FO_PathSplit
; Описание ........: Делит путь на 3 части: путь, файл, расширение.
; Синтаксис.......: _FO_PathSplit ( $sPath )
; Параметры:
;		$sPath - Путь
; Возвращаемое значение: Успешно - массив из 3-x элементов следующего формата
;		$Array[0] = путь
;		$Array[1] = имя файла / каталога
;		$Array[2] = расширение
; Автор ..........: AZJIO
; Примечания ..: Функция не имеет ошибок. Если отсутствует какой либо элемент пути, то массив содержит пустую ячейку для этого элемента
; ============================================================================================
Func _FO_PathSplit($sPath)
	Local $i, $aPath[3] ; ( Dir | Name | Ext )
	$i = StringInStr($sPath, '\', 0, -1)
	$aPath[1] = StringTrimLeft($sPath, $i)
	$aPath[0] = StringLeft($sPath, $i) ; Dir
	$i = StringInStr($aPath[1], '.', 1, -1)
	If $i Then
		$i -= 1
		$aPath[2] = StringTrimLeft($aPath[1], $i) ; Ext
		$aPath[1] = StringLeft($aPath[1], $i) ; Name
	EndIf
	Return $aPath
EndFunc   ;==>_FO_PathSplit

Func _FO_IsDir($sTmp)
	$sTmp = FileGetAttrib($sTmp)
	Return SetError(@error, 0, StringInStr($sTmp, 'D', 2) > 0)
EndFunc   ;==>_FO_IsDir

; #FUNCTION# ;=================================================================================
; Function Name ...: _FO_ShortFileSize
; Description ........: The file size in bytes converts into short to 3 significant figures
; Syntax................: _FO_ShortFileSize($iBytes)
; Parameters:
;		$iBytes - Bytes
; Return values ....: A string containing the number
; Author(s) ..........: AZJIO
; Remarks ..........:
; ============================================================================================
; Имя функции ...: _FO_ShortFileSize
; Описание ........: Преобразует байты в число, округлённое до 3 знаков
; Синтаксис.......: _FO_ShortFileSize($iBytes)
; Параметры:
;		$iBytes - Число байт
; Возвращаемое значение: Строка, содержащая число с приставкой
; Автор ..........: AZJIO
; Примечания ..:
; ============================================================================================
Func _FO_ShortFileSize($iBytes)
	Switch $iBytes
		Case 10995116277760 To 109951162777600 ; 10 - 100 TB
			$iBytes = Round($iBytes / 1099511627776, 1) & ' TB'
		Case 1000000000000 To 10995116277759 ; 1000 GB - 10 TB
			$iBytes = Round($iBytes / 1099511627776, 2) & ' TB'
		Case 107374182400 To 999999999999 ; 100 - 999 GB
			$iBytes = Round($iBytes / 1073741824) & ' GB'
		Case 10737418240 To 107374182399 ; 10 - 100 GB
			$iBytes = Round($iBytes / 1073741824, 1) & ' GB'
		Case 1000000000 To 10737418239 ; 1000 MB - 10 GB
			$iBytes = Round($iBytes / 1073741824, 2) & ' GB'
		Case 104857600 To 999999999 ; 100 - 999 MB
			$iBytes = Round($iBytes / 1048576) & ' MB'
		Case 10485760 To 104857599 ; 10 - 100 MB
			$iBytes = Round($iBytes / 1048576, 1) & ' MB'
		Case 1000000 To 10485759 ; 1000 KB - 10 MB
			$iBytes = Round($iBytes / 1048576, 2) & ' MB'
		Case 102400 To 999999 ; 100 - 999 KB
			$iBytes = Round($iBytes / 1024) & ' KB'
		Case 10240 To 102399 ; 10 - 100 KB
			$iBytes = Round($iBytes / 1024, 1) & ' KB'
		Case 1000 To 10239 ; 1000 B - 10 KB
			$iBytes = Round($iBytes / 1024, 2) & ' KB'
		Case 0 To 999
			$iBytes &= ' B'
	EndSwitch
	Return $iBytes
EndFunc   ;==>_FO_ShortFileSize

; #FUNCTION# ;=================================================================================
; Function Name ...: _FO_IsEmptyFolder
; Description ........: Checks whether the folder is empty
; Syntax................: _FO_IsEmptyFolder($sPath)
; Parameters:
;		$sPath - The path to the folder
; Return values ....: Success
;					|1 - the folder is empty
;					|0 - folder is not empty or error.
;					Failure - If @error is set to 1, the folder does not exist or access denied.
; Author(s) ..........: AZJIO
; Remarks ..........:
; ============================================================================================
; Имя функции ...: _FO_IsEmptyFolder
; Описание ........: Проверяет, является ли папка пустой
; Синтаксис.......: _FO_IsEmptyFolder($sPath)
; Параметры:
;		$sPath - Путь к папке
; Возвращаемое значение: Успешно
;					|1 - папка пуста
;					|0 - папка не пуста или ошибка.
;					Неудачно - Если @error равен 1, то папка не существует или нет доступа.
; Автор ..........: AZJIO
; Примечания ..:
; ============================================================================================
Func _FO_IsEmptyFolder($sPath)
	If StringRight($sPath, 1) <> '\' Then $sPath &= '\'
	Local $s = FileFindFirstFile($sPath & '*')
	If $s = -1 Then
		If @error Then Return 1
		Return SetError(1, 0, 0)
	Else
		FileClose($s)
		Return 0
	EndIf
EndFunc   ;==>_FO_IsEmptyFolder

; #FUNCTION# ;=================================================================================
; Function Name ...: _FO_CreateFile
; Description ........: Creates the file
; Syntax................: _FO_CreateFile($sPath [, $sData = '' [, $flag = 0 ]] )
; Parameters:
;		$sPath - The path to the folder
;		$sData - The data to be written to the file
;		$flag - If not 0, it overwrites the existing file
; Return values ....:
;					|Success -  Returns 1 - file exists or is created
;					|Failed - Returns 0 - the file does not exist and is not created
; Author(s) ..........: AZJIO
; Remarks ..........:
; ============================================================================================
; Имя функции ...: _FO_CreateFile
; Описание ........: Создаёт файл
; Синтаксис.......: _FO_CreateFile($sPath [, $sData = '' [, $flag = 0 ]] )
; Параметры:
;		$sPath - Путь к папке
;		$sData - Данные для записи в файл
;		$flag - Если не 0, то перезаписывает существующий файл
; Возвращаемое значение:
;					|Успешно -  Возвращает 1 - файл существовал или создан
;					|Неудачно - Возвращает 0 - файл не существовал и не создан
; Автор ..........: AZJIO
; Примечания ..:
; ============================================================================================
Func _FO_CreateFile($sPath, $sData = '', $flag = 0)
	If Not FileExists($sPath) Or $flag Then
		Local $hFile = FileOpen($sPath, 10)
		Local $iRet = FileWrite($hFile, $sData)
		FileClose($hFile)
		Return $iRet
	EndIf
	Return 1
EndFunc   ;==>_FO_CreateFile