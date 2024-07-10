@Echo Starting Charge Correction File Combine

SETLOCAL
set FILE_PATH=%~dp0
set SCRIPT_PATH=%FILE_PATH%input_main.py
python -u "%SCRIPT_PATH%"
ENDLOCAL

@Echo Process Completed
pause