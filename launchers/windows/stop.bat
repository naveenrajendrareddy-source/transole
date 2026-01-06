@echo off
TITLE Stop Transol VMS
echo =====================================================
echo      Stop Transol VMS
echo =====================================================
echo.

set FOUND=0
FOR /F "tokens=5" %%P IN ('netstat -a -n -o ^| findstr ":8000" ^| findstr /i "LISTENING"') DO (
    echo [INFO] Found server running (PID: %%P). Stopping...
    taskkill /F /PID %%P
    set FOUND=1
)

IF %FOUND%==0 (
    echo [INFO] No server found running on port 8000.
) ELSE (
    echo [SUCCESS] Server stopped.
)

echo.
pause
