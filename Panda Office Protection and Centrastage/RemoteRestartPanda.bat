@ECHO OFF
if not '%1'=='' goto %1

:START
echo This tool will remotely restart the Panda Systems Management Agent
ECHO.
echo  4G Wireless IT Department
ECHO __________________
ECHO.
Echo 1.- Device is on the WAN
Echo 0.- Device is on the LAN
ECHO.
ECHO.
ECHO.

set choice=
set /p choice=Type the option: 
if not '%choice%'=='' set choice=%choice:~0,1%


if '%choice%'=='0' goto DOMAINORNAH
if '%choice%'=='1' goto ONWAN
ECHO "%choice%" Character invalid , type it again please.
goto START


:DOMAINORNAH
set /p DOMAIN=Is this on the domain? Yes or no: 

IF "%DOMAIN%"== "yes"  (
		GOTO ONDOMAIN
) ELSE (
		GOTO NOTDOMAIN
)

:ONWAN
set /p IPADDR=IP Address:
set /p USERNM=Admin username:

EditV64.exe -m -p "Enter password " PASSWORD
wmic /node:%IPADDR% /user:%USERNM% /password:%Password% service where caption="Panda Cloud Systems Management" call stopservice
wmic /node:%IPADDR% /user:%USERNM% /password:%Password% service where caption="Panda Cloud Systems Management" call startservice

PAUSE

GOTO ONEMORE

:NOTDOMAIN
set /p COMPUTERNAME=Hostname:
set /p USERNM=Admin username: 

EditV64.exe -m -p "Enter password " PASSWORD

for /f "tokens=2 delims=[]" %%f in ('ping -4 -n 1 %COMPUTERNAME% ^| find /i "pinging"') do set "IPADDR=%%~f"
echo IP: %IPADDR%

wmic /node:%IPADDR% /user:%USERNM% /password:%Password% service where caption="Panda Cloud Systems Management" call stopservice
wmic /node:%IPADDR% /user:%USERNM% /password:%Password% service where caption="Panda Cloud Systems Management" call startservice

PAUSE

GOTO ONEMORE

:ONDOMAIN
set /p COMPUTERNAME=Hostname: 
set /p USERNM=domain\username: 

EditV64.exe -m -p "Enter password " PASSWORD

for /f "tokens=2 delims=[]" %%f in ('ping -4 -n 1 %COMPUTERNAME% ^| find /i "pinging"') do set "IPADDR=%%~f"
echo IP: %IPADDR%

wmic /node:%IPADDR% /user:%USERNM% /password:%Password% service where caption="Panda Cloud Systems Management" call stopservice
wmic /node:%IPADDR% /user:%USERNM% /password:%Password% service where caption="Panda Cloud Systems Management" call startservice

PAUSE

GOTO ONEMORE

:ONEMORE
cls
ECHO Last PC service was restarted on
nbtstat -A %IPADDR%
ECHO.
ECHO.
ECHO Would you like to restart the service on another computer?
ECHO.
Echo 1.- YES
Echo 0.- NO
set choice=
set /p choice=Type the option: 
if not '%choice%'=='' set choice=%choice:~0,1%


if '%choice%'=='0' goto END
if '%choice%'=='1' goto START
ECHO "%choice%" Character invalid , type it again please.
goto START

:END
echo Press any key to exit...
pause >nul