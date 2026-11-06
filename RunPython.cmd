@echo off
py -3.12 -c "import sys; sys.exit(sys.version_info[:2] != (3,12))" >nul 2>nul
if not errorlevel 1 goto launcher
python -c "import sys; sys.exit(sys.version_info[:2] != (3,12))" >nul 2>nul
if not errorlevel 1 goto python
echo Python 3.12 was not found. Check your Python installation or PATH.
exit /b 1
:launcher
py -3.12 %*
exit /b %errorlevel%
:python
python %*
exit /b %errorlevel%
