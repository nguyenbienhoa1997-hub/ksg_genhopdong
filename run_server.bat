@echo off
title KSG HopDong Server
cd /d "%~dp0"

:: Tu dong lay IP LAN
for /f "tokens=2 delims=:" %%a in ('ipconfig ^| findstr /R /C:"IPv4.*10\." /C:"IPv4.*192\." /C:"IPv4.*172\."') do (
    set LAN_IP=%%a
    goto :found
)
:found
set LAN_IP=%LAN_IP: =%

echo.
echo  ==========================================
echo   KSG Hop Dong - LAN Server
echo   http://%LAN_IP%:5000
echo  ==========================================
echo.
venv\Scripts\python.exe -m waitress --host=0.0.0.0 --port=5000 --threads=16 --connection-limit=100 --channel-timeout=60 app:app
pause
