@Echo Starting Charge Correction File Review

SETLOCAL
set FILE_PATH=%~dp0
set SCRIPT_PATH=%FILE_PATH%FileReview.py
python -u "%SCRIPT_PATH%"
ENDLOCAL

@Echo Process Completed
pause