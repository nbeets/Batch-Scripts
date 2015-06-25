@ECHO OFF

PUSHD %~dp0

echo This is remotely restart the Panda Cloud Systems Management service

:WANORLAN
set /p NETWORK=Is this on the LAN or WAN? WAN or LAN:

IF "%NETWORK%"== "wan"  (
		GOTO ONWAN
) ELSE (
		GOTO DOMAINORNAH
)

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
set /p ONEMORE=Restart service on another endpoint? YES or NO:

IF "%ONEMORE%"== "yes"  (
		GOTO WANORLAN
) ELSE (
		GOTO END
)


:END
echo Press any key to exit...
pause >nul