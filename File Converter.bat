@echo off
rem Launch File Converter without a console window.
if not exist "%~dp0.venv\Scripts\pythonw.exe" (
  echo First run: installing dependencies...
  call "%~dp0setup.bat" || exit /b 1
)
start "" "%~dp0.venv\Scripts\pythonw.exe" "%~dp0app.py"
