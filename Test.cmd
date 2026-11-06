@echo off
setlocal
call "%~dp0RunPython.cmd" "%~dp0bootstrap.py" test
set "search_exit=%errorlevel%"
echo.
pause
exit /b %search_exit%
