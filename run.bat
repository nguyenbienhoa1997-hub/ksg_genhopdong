@echo off
chcp 65001 > nul
title Tao Hop Dong PDF

echo.
echo  ============================================
echo     TAO HOP DONG PDF - Contract Generator
echo  ============================================
echo.

:: Check Python
python --version > nul 2>&1
if errorlevel 1 (
    echo [LOI] Khong tim thay Python.
    echo Vui long cai Python 3.8+ tu https://python.org
    pause
    exit /b 1
)

:: Create venv if needed
if not exist venv (
    echo [1/3] Tao virtual environment...
    python -m venv venv
)

:: Activate
call venv\Scripts\activate.bat

:: Install deps if flag file missing
if not exist venv\.installed (
    echo [2/3] Cai dat thu vien (lan dau, mat vai phut)...
    pip install -r requirements.txt -q
    echo. > venv\.installed
) else (
    echo [2/3] Thu vien da san sang.
)

echo [3/3] Khoi dong server...
echo.
echo  >> Trinh duyet se tu mo tai: http://localhost:5000
echo  >> Nhan Ctrl+C de dung server.
echo.

start "" "http://localhost:5000"
python app.py

pause
