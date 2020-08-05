@ECHO OFF

IF %1.==. GOTO START
IF %1.==search. GOTO SEARCH
IF %1.==file. GOTO FILE
GOTO END

:START
TASKKILL /IM XXX.EXE /f
TASKKILL /IM XXXXX.EXE /f
CALL %0 search "C:\XXX"
GOTO END

:SEARCH
IF %2.==. GOTO ERROR
echo %2
FOR /R %2 %%i IN (*.*) DO CALL %0 file "%%i"
GOTO END

:FILE
IF %2.==. GOTO ERROR
IF %2=="C:\XXX\XXX.EXE" GOTO NODEL
IF %2=="C:\XXX\XXXXXXX.DLL" GOTO NODEL
ECHO Удаление файла: %2
DEL %2 /Q /F
GOTO END

:NODEL
ECHO Оставляем файл: %2
GOTO END

:ERROR
ECHO Ошибка!

:END