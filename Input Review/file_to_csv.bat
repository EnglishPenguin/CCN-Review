@Echo Saving the Input File as a .csv in the correct location

SETLOCAL
set FILE_PATH=%~dp0
set SCRIPT_PATH=%FILE_PATH%file_to_csv.py
python -u "%SCRIPT_PATH%"
ENDLOCAL

@Echo Process Completed