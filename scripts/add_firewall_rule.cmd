@echo off
set PORT=8765

netsh advfirewall firewall show rule name="Hymn Search Remote" >nul 2>&1
if %errorlevel%==0 goto exists

netsh advfirewall firewall add rule name="Hymn Search Remote" dir=in action=allow protocol=TCP localport=%PORT% profile=private,domain,public enable=yes
if %errorlevel% neq 0 goto fail

echo OK: firewall rule added for TCP %PORT%
echo Test from another PC: http://YOUR_PC_IP:%PORT%/api/health
goto end

:exists
echo OK: rule already exists
netsh advfirewall firewall show rule name="Hymn Search Remote"
goto end

:fail
echo ERROR: Right-click this file and choose "Run as administrator"

:end
pause
