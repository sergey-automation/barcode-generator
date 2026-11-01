@echo off
setlocal
cd /d "%~dp0"

set PYEXE=

where py >nul 2>nul
if not errorlevel 1 set PYEXE=py

if not defined PYEXE (
    where python >nul 2>nul
    if not errorlevel 1 set PYEXE=python
)

if not defined PYEXE (
    echo Python 3 was not found.
    echo Install Python 3 and run BUILD_EXE.bat again.
    pause
    exit /b 1
)

echo Installing/updating PyInstaller...
%PYEXE% -m pip install --user --upgrade pyinstaller
if errorlevel 1 (
    echo PyInstaller installation failed.
    pause
    exit /b 1
)

echo.
echo Building standalone Windows executable...
%PYEXE% -m PyInstaller --noconfirm --clean --onefile --console --name barcode_generator barcode_generator.py
if errorlevel 1 (
    echo Build failed.
    pause
    exit /b 1
)

copy /Y "dist\barcode_generator.exe" "..\release_ru\barcode_generator.exe" >nul
copy /Y "dist\barcode_generator.exe" "..\release_en\barcode_generator.exe" >nul

echo.
echo Creating release ZIP files...
powershell -NoProfile -Command "Compress-Archive -Force -Path '..\release_ru\barcode_generator.exe','..\release_ru\barcode_generator.xlsx','..\release_ru\REFERENCE_RU.txt' -DestinationPath '..\ean13-barcode-generator-RU-release.zip'"
powershell -NoProfile -Command "Compress-Archive -Force -Path '..\release_en\barcode_generator.exe','..\release_en\barcode_generator.xlsx','..\release_en\REFERENCE_EN.txt' -DestinationPath '..\ean13-barcode-generator-EN-release.zip'"

echo.
echo DONE
echo EXE copied to release_ru and release_en.
echo Release ZIP files created in the repository root.
pause
