@ECHO OFF
PUSHD %~dp0

taskkill /f /im iTunes.exe /fi "memusage gt 40" 2>NUL | findstr SUCCESS >NUL && if errorlevel 1 ( echo iTunes was not killed ) else ( echo iTunes was killed )

wmic product where vendor="Apple Inc." call uninstall

start /wait msiexec ALLUSERS=true reboot=suppress /qn /i "AppleApplicationSupport64.msi" /log %SystemDrive%\install.log
start /wait msiexec ALLUSERS=true reboot=suppress /qn /i "AppleSoftwareUpdate.msi" /log %SystemDrive%\install.log
start /wait msiexec /qn /norestart /i "AppleMobileDeviceSupport6464" /log %SystemDrive%\install.log
start /wait msiexec /qn /norestart /i "Bonjour64.msi" /log %SystemDrive%\install.log
start /wait msiexec /qn /norestart /i "iTunes6464.msi"  /log %SystemDrive%\install.log