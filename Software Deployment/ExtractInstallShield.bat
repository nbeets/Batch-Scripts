@ECHO OFF

PUSHD %~dp0

mkdir extracted

set /p EXENAME=Name of exe: 

%EXENAME% /s /x /b"extracted" /v"/qn"