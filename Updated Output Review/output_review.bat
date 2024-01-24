@Echo off

SETLOCAL
set FILE_PATH=%~dp0
set SCRIPT_PATH=%FILE_PATH%output_main.py
python -u "%SCRIPT_PATH%"
ENDLOCAL