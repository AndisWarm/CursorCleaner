@echo off
setlocal EnableExtensions
rem Always run the script beside this BAT file.
set "ROOT=%~dp0"
cd /d "%ROOT%"
set "SCRIPT=%ROOT%cursor_cleaner.py"
if not exist "%SCRIPT%" (
    echo [ERROR] Script not found: "%SCRIPT%"
    pause
    exit /b 2
)
for %%I in ("%SCRIPT%") do echo [Cursor Cleaner] Using %%~fI ^(modified %%~tI^)

rem Use UTF-8 output.
chcp 65001 >nul
set "PYTHONUTF8=1"
set "PYTHONIOENCODING=utf-8"

rem Prefer the project virtual environment and always execute the current .py file.
if exist "%ROOT%.venv\Scripts\python.exe" goto run_venv
where python >nul 2>&1
if not errorlevel 1 goto run_path_python
where py >nul 2>&1
if not errorlevel 1 goto run_py_launcher
echo [ERROR] Python 3.10+ was not found.
pause
exit /b 3

:run_venv
"%ROOT%.venv\Scripts\python.exe" -B -X utf8 "%SCRIPT%" %*
goto finish

:run_path_python
python -B -X utf8 "%SCRIPT%" %*
goto finish

:run_py_launcher
py -3 -B -X utf8 "%SCRIPT%" %*

:finish
set "exit_code=%ERRORLEVEL%"
endlocal & exit /b %exit_code%
