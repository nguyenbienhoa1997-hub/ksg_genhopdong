@echo off
:: Xin quyen Admin tu dong
net session >nul 2>&1
if %errorLevel% neq 0 (
    powershell -Command "Start-Process '%~f0' -Verb RunAs"
    exit /b
)

netsh advfirewall firewall add rule name="KSG HopDong Port 5000" dir=in action=allow protocol=TCP localport=5000
echo.
echo  Da mo port 5000 thanh cong!
echo  May khac co the vao: http://10.62.4.147:5000
echo.
pause
