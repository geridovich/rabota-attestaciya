@echo off
rem Сборка dist\AttestIK.exe (один файл, Python на рабочем ПК не нужен).
rem Собирается 32-битным Python 3.8: такой exe работает на Windows 7/8/10/11,
rem и 32-, и 64-битных. Путь к 32-битному Python можно передать первым параметром.
cd /d "%~dp0"
set PY32=%1
if "%PY32%"=="" set PY32=py -3.8-32
if not exist .venv32 %PY32% -m venv .venv32
.venv32\Scripts\python -m pip install -r requirements.txt pyinstaller
.venv32\Scripts\pyinstaller --noconfirm --onefile --windowed --name AttestIK main.py
