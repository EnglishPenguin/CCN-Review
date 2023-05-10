@Echo Preparing the Charge Correction Email

SETLOCAL
set FILE_PATH=%~dp0
set SCRIPT_PATH=%FILE_PATH%EmailPrep.py
python -u "%SCRIPT_PATH%"
ENDLOCAL

@Echo Process Completed