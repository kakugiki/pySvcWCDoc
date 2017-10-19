ECHO OFF

call "%VS140COMNTOOLS%\vsvars32.bat"

REM check if PySvc exists
sc query PySvc | find "does not exist" > nul
if %ERRORLEVEL% EQU 1 GOTO EXISTS
if %ERRORLEVEL% EQU 0 GOTO MISSING

:EXISTS
REM sc stop PySvc
sc stop PySvc

REM sc delete PySvc
sc delete PySvc

:MISSING

REM C:\Users\wguo\PycharmProjects\pySvcWCDoc
CD %~dp0

python.exe .\pySvcWCDoc.py install

net start PySvc 

pause