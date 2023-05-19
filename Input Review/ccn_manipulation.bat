@Echo Starting Charge Correction File Manipulation

SETLOCAL
set FILE_PATH=%~dp0
set SCRIPT_PATH=%FILE_PATH%FileManipulation.py
python -u "%SCRIPT_PATH%"
ENDLOCAL

@Echo Process Completed
pause