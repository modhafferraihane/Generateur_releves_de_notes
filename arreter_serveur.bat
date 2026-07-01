@echo off
echo Arret du serveur Flask sur le port 5000...
for /f "tokens=5" %%a in ('netstat -aon ^| findstr ":5000 " ^| findstr "LISTENING"') do (
    taskkill /PID %%a /F >nul 2>&1
    echo Serveur arrete (PID %%a).
)
echo Fait.
pause
