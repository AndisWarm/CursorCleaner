@echo off
setlocal EnableExtensions
rem Codex/PowerShell 可能仍使用 GBK；先把当前 cmd 会话切到 UTF-8。
chcp 65001 >nul
set "PYTHONUTF8=1"
set "PYTHONIOENCODING=utf-8"
python -X utf8 "%~dp0cursor_cleaner.py" %*
set "exit_code=%ERRORLEVEL%"
endlocal & exit /b %exit_code%
