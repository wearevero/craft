@echo off

:: Path to Python
set PYTHON=C:\Python313\python.exe

:: Path to backup.py
set SCRIPT=backup.py

:: Path to log file
set LOGFILE=D:\python_auto_run.log

echo === %date% %time% === >> "%LOGFILE%"
"%PYTHON%" "%SCRIPT%" >> "%LOGFILE%" 2>&1
echo. >> "%LOGFILE%"
