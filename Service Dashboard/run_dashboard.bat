@echo off
REM Double-click after dropping the files into the Dropbox folder.
set FOLDER=%USERPROFILE%\Dropbox\GoAxen Drop
python "%~dp0goaxen_dashboard.py" "%FOLDER%"
start "" "%FOLDER%\_output\dashboard.html"
pause
