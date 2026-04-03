@echo off
cd /d C:\Users\Jeremy\digital-invoice-organizer
echo Stopping existing Node processes...
taskkill /F /IM node.exe >nul 2>nul
timeout /t 1 /nobreak >nul
echo Restarting digital-invoice-organizer...
start "digital-invoice-organizer" cmd /k "cd /d C:\Users\Jeremy\digital-invoice-organizer && npm start"
