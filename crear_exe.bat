@echo off
setlocal ENABLEDELAYEDEXPANSION
title Build BioAccess (Debug/Console)

set "NAME=BioAccess_Debug"
set "PY=.\.venv\Scripts\python.exe"
if not exist ".\.venv\Scripts\python.exe" (
  echo [!] No se encontro .venv\Scripts\python.exe. Usare 'python' del sistema.
  set "PY=python"
)

rmdir /s /q build  2>nul
rmdir /s /q dist   2>nul
del /q "%NAME%.spec" 2>nul

%PY% -m pip install -U pip setuptools wheel pyinstaller

rem -- Construye add-data solo si existen las carpetas
set "ADD_DATA="
for %%D in ("assets" "face_recognition_models" "rostros" "salidas_solicitudes" "formularios") do (
  if exist "%%~fD" set "ADD_DATA=!ADD_DATA! --add-data ""%%~fD;%%~nxD"""
)
echo === ADD_DATA: !ADD_DATA!

%PY% -m PyInstaller ^
  --noconfirm ^
  --clean ^
  --name "%NAME%" ^
  --console ^
  --icon assets\bioaccess.ico ^
  --hidden-import=pkg_resources ^
  --hidden-import=asistencia_funcionarios ^
  --hidden-import=asistencia_diaria ^
  --collect-all customtkinter ^
  --collect-all tkcalendar ^
  --collect-all face_recognition ^
  --collect-all dlib ^
  --collect-all cv2 ^
  --collect-all PIL ^
  --collect-all pandas ^
  --collect-all numpy ^
  --collect-all openpyxl ^
  --collect-all xlsxwriter ^
  --collect-all reportlab ^
  --collect-data face_recognition_models ^
  --collect-data tkcalendar ^
  --collect-binaries cv2 ^
  !ADD_DATA! ^
  principal.py

if errorlevel 1 (
  echo ❌ PyInstaller fallo. Revisa el log.
  pause & exit /b 1
)

if exist reloj_control.db copy /Y reloj_control.db "dist\%NAME%\reloj_control.db" >nul

for %%D in ("rostros" "salidas_solicitudes") do (
  if not exist "dist\%NAME%\%%~nxD" mkdir "dist\%NAME%\%%~nxD"
)

echo(
echo ✅ Debug listo con consola.
echo ▶️ Ejecutable: dist\%NAME%\%NAME%.exe
pause
