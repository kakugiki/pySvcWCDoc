ECHO OFF

call "%VS140COMNTOOLS%\vsvars32.bat"

REM check if PySvcWCDoc exists
sc query PySvcWCDoc | find "does not exist" > nul
if %ERRORLEVEL% EQU 1 GOTO EXISTS
if %ERRORLEVEL% EQU 0 GOTO MISSING

:EXISTS
REM sc stop PySvcWCDoc
sc stop PySvcWCDoc

REM sc delete PySvcWCDoc
sc delete PySvcWCDoc

:MISSING

REM C:\Users\wguo\PycharmProjects\PySvcWCDoc
CD %~dp0

python.exe .\PySvcWCDoc.py --startup=delayed install

net start PySvcWCDoc 

pause