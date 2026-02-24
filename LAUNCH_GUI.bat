@echo off
REM Quick launcher for Certificate Generator GUI
echo Starting Certificate Generator...
cd /d "%~dp0.."
.venv\Scripts\python.exe GUI_Application\gui_certificate_generator.py
pause
