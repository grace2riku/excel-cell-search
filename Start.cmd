@echo off
setlocal
call "%~dp0RunPython.cmd" "%~dp0bootstrap.py" start %*
set "search_exit=%errorlevel%"
echo.
pause
exit /b %search_exit%
