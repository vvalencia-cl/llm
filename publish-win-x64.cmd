@echo off
setlocal

REM Publish portable single-file build (win-x64, self-contained)
dotnet publish LMM\LMM.csproj -c Release -r win-x64 --self-contained true ^
  /p:PublishSingleFile=true ^
  /p:IncludeNativeLibrariesForSelfExtract=true ^
  /p:EnableCompressionInSingleFile=true

IF ERRORLEVEL 1 (
  echo.
  echo Publish FAILED.
  exit /b 1
)

set PUB_DIR=LMM\bin\Release\net10.0-windows\win-x64\publish

REM Extract version from .csproj
for /f "usebackq tokens=*" %%v in (`powershell -NoProfile -Command "$xml=[xml](Get-Content 'LMM\LMM.csproj'); $xml.Project.PropertyGroup.Version | Where-Object {$_} | Select-Object -First 1"`) do set APP_VERSION=%%v

if not "%APP_VERSION%"=="" (
    echo.
    echo Version detected: %APP_VERSION%
    echo Renaming LMM.exe to LMM_%APP_VERSION%.exe
    move /y "%PUB_DIR%\LMM.exe" "%PUB_DIR%\LMM_%APP_VERSION%.exe" >nul
    set EXE_NAME=LMM_%APP_VERSION%.exe
) else (
    set EXE_NAME=LMM.exe
)

echo.
echo Publish OK.
echo Output: %PUB_DIR%\
echo Executable: %PUB_DIR%\%EXE_NAME%
endlocal
pause
