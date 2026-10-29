@echo off
rem Сборка dist\AttestIK.exe (один файл, Python на рабочем ПК не нужен).
rem Собирается 32-битным Python 3.7 (последний с поддержкой Vista): exe работает
rem на Windows Vista/7/8/10/11, 32- и 64-битных.
rem   %1 - команда 32-битного Python 3.7 (по умолчанию py -3.7-32)
rem   %2 - папка с app-local UCRT (ucrtbase.dll и api-ms-win-*.dll, 32 бит) для
rem        Vista/7 без обновления KB2999226, например bin 32-битной Java 8
cd /d "%~dp0"
set PY=%~1
if "%PY%"=="" set PY=py -3.7-32
set UCRT=%~2
if "%UCRT%"=="" set UCRT=C:\Program Files (x86)\Eclipse Adoptium\jre-8.0.472.8-hotspot\bin
if not exist .venv37 %PY% -m venv .venv37
.venv37\Scripts\python -m pip install -r requirements-legacy.txt "lxml<4.10" "pyinstaller<6"
set EXTRA=
if exist "%UCRT%\ucrtbase.dll" (
  for %%f in ("%UCRT%\api-ms-win-*.dll" "%UCRT%\ucrtbase.dll") do call set EXTRA=%%EXTRA%% --add-binary "%%~f;."
)
.venv37\Scripts\pyinstaller --noconfirm --onefile --windowed --name AttestIK %EXTRA% main.py
