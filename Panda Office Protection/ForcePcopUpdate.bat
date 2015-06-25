@ECHO OFF

REM x86.
IF /I '%PROCESSOR_ARCHITECTURE%' == 'x86' GOTO 32BIT

REM x64, AMD64, EM64T.
IF /I '%PROCESSOR_ARCHITECTURE%' == 'AMD64' GOTO 64BIT




:64BIT
net stop wahost
TASKKILL /F /IM WAHost.exe
C:
cd "%programfiles(x86)%\Panda Security\WaAgent\Waltest"
del *.dat
cd "%programfiles(x86)%\Panda Security\WaAgent\WalUpd"
del /S /Q Data
Net start wahost
cd "%programfiles(x86)%\Panda Security\WaAgent\WasLpMng"
GOTO fix2


:32BIT
net stop wahost
TASKKILL /F /IM WAHost.exe
C:
cd "%programfiles%\Panda Security\WaAgent\Waltest"
del *.dat
cd "%programfiles%\Panda Security\WaAgent\WalUpd"
del /S /Q Data
Net start wahost
cd "%programfiles%\Panda Security\WaAgent\WasLpMng"
GOTO fix2


:fix2

WAPLPMNG.exe walforce -force -realtime

WAPLPMNG.exe walupd -force

WAPLPMNG.exe waltest -force

:END