@echo off
title KSG HopDong Server
cd /d "%~dp0"
echo.
echo  ==========================================
echo   KSG Hop Dong - LAN Server
echo   http://11.20.10.53:5000
echo  ==========================================
echo.
venv\Scripts\python.exe -m waitress --host=0.0.0.0 --port=5000 --threads=16 --connection-limit=100 --channel-timeout=60 app:app
pause
