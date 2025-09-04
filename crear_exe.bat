@echo off
setlocal

set PROY=C:\Proyectos\RelojControl
set BUILD=C:\Proyectos\RelojControl_build_console
set DIST=%BUILD%\dist
set RELEASE=C:\Proyectos\RelojControl_release_console

if exist "%BUILD%" rmdir /s /q "%BUILD%"
mkdir "%BUILD%"

cd /d "%PROY%"

pyinstaller ^
 --noconfirm ^
 --console ^
 --onefile ^
 --name BioAccessConsole ^
 --icon "%PROY%\assets\bioaccess.ico" ^
 --distpath "%DIST%" ^
 --workpath "%BUILD%\build" ^
 --specpath "%BUILD%" ^
 --add-data "%PROY%\face_recognition_models;face_recognition_models" ^
 --collect-data tkcalendar ^
 --collect-data customtkinter ^
 --collect-submodules cv2 ^
 "%PROY%\principal.py"
if errorlevel 1 ( echo ERROR de compilacion & pause & exit /b 1 )

if exist "%RELEASE%" rmdir /s /q "%RELEASE%"
mkdir "%RELEASE%"

copy "%DIST%\BioAccessConsole.exe" "%RELEASE%\BioAccessConsole.exe" >nul

REM Copia DB y carpetas que tu app usa/crea
if exist "%PROY%\reloj_control.db" copy "%PROY%\reloj_control.db" "%RELEASE%\reloj_control.db" >nul
if exist "%PROY%\formularios" xcopy "%PROY%\formularios" "%RELEASE%\formularios" /E /I /Y >nul
mkdir "%RELEASE%\rostros"
mkdir "%RELEASE%\salidas_solicitudes"
mkdir "%RELEASE%\exportes_horario"

REM Copiar DLL de video de OpenCV si existe (ajusta ruta si difiere)
for %%D in ("%USERPROFILE%\AppData\Local\Programs\Python\Python3*\Lib\site-packages\cv2\opencv_videoio_ffmpeg*.dll") do (
  copy "%%~fD" "%RELEASE%\" >nul
)

echo Listo: %RELEASE%\BioAccessConsole.exe
pause
endlocal
