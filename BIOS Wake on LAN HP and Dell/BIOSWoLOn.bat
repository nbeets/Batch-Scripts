@ECHO OFF

:: Force working directory to be current directory
PUSHD %~dp0

:: Gets vendor name from WMIC and shortens output into a variable a
FOR /F "tokens=2 delims==" %%A IN ('WMIC csproduct GET vendor /VALUE ^| FIND /I "Vendor="') DO SET machine=%%A

:: Check if vendor is Dell
:CHECK1
IF "%machine%"== "Dell Inc." (
	GOTO DELL1
) ELSE ( 
	GOTO CHECK2
)

:: Check if vendor is HP
:CHECK2
IF "%machine%"== "Hewlett-Packard" (
	GOTO HP
) ELSE (
	GOTO NEITHER
)

:: Runes Dell's cctk utility to turn on WOL
:DELL1
IF EXIST "C:\Program Files (x86)\Dell\Command Configure\x86_64\" (
	GOTO 64BITDELL
) ELSE ( 
	GOTO DELL2
)

:DELL2
IF EXIST "C:\Program Files\Dell\Command Configure\x86_64\" (
	GOTO 32BITDELL
) ELSE ( 
	GOTO DELLINSTALL
)

:DELLINSTALL
start /wait msiexec /i Command_Configure.msi /qn

IF /I '%PROCESSOR_ARCHITECTURE%' == 'x86' GOTO 32BITDELL

IF /I '%PROCESSOR_ARCHITECTURE%' == 'AMD64' GOTO 64BITDELL

:64BITDELL
cd "C:\Program Files (x86)\Dell\Command Configure\x86_64"
cctk.exe --wakeonlan=enable

GOTO END

:32BITDELL
cd "C:\Program Files\Dell\Command Configure\x86"
cctk.exe --wakeonlan=enable

GOTO END

:HP
::Checks for CPU architecture since utility is different
IF /I '%PROCESSOR_ARCHITECTURE%' == 'x86' GOTO 32BIT

IF /I '%PROCESSOR_ARCHITECTURE%' == 'AMD64' GOTO 64BIT

:: Runs HP's 64-bit BIOS configuration utility
:64BIT
biosconfigutility64.exe /Set:"config.txt"

GOTO END

:: Runs HP's 32-bit BIOS configuration utility
:32BIT
biosconfigutility.exe /Set:"config.txt"

GOTO END

:NEITHER
ECHO THIS IS NOT HP OR DELL
GOTO END

:END