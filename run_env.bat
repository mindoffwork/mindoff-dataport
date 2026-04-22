@echo off
REM Exit on error
IF ERRORLEVEL 1 EXIT /B

REM Activate virtual environment
CALL .venv\Scripts\activate.bat
ECHO ✅ Virtual environment activated
