@echo off
rem Сборка одного файла AttestIK.exe в папке dist (Python на рабочем ПК не нужен)
cd /d "%~dp0"
if not exist .venv py -3 -m venv .venv
.venv\Scripts\python -m pip install -r requirements.txt pyinstaller
.venv\Scripts\pyinstaller --noconfirm --onefile --windowed --name AttestIK main.py
