@echo off
rem Сборка dist\PlusUltra64.exe - 64-битная версия для Windows 10/11.
rem Ей доступна вся память компьютера (32-битной - не больше ~2 ГБ), и она быстрее.
rem Для Vista/7 и 32-битной Windows по-прежнему нужен PlusUltra.exe (build_exe.bat).
rem   %1 - команда 64-битного Python 3.8+ (по умолчанию py -3)
cd /d "%~dp0"
set PY=%~1
if "%PY%"=="" set PY=py -3
if not exist .venv %PY% -m venv .venv
.venv\Scripts\python -m pip install -r requirements.txt pyinstaller
rem модули для разработки (картинки помощи, звук, PDF) в программу не нужны
.venv\Scripts\pyinstaller --noconfirm --onefile --windowed --name PlusUltra64 --exclude-module numpy --exclude-module PIL --exclude-module soundfile --exclude-module lameenc --exclude-module pypdfium2 --exclude-module cffi --icon attest\assets\app.ico --add-data "attest\templates;attest\templates" --add-data "attest\assets;attest\assets" main.py
