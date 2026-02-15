@echo off
setlocal

REM Publish portable single-file build (win-x64, self-contained)
dotnet publish -c Release -r win-x64 --self-contained true ^
  /p:PublishSingleFile=true ^
  /p:IncludeNativeLibrariesForSelfExtract=true ^
  /p:EnableCompressionInSingleFile=true

IF ERRORLEVEL 1 (
  echo.
  echo Publish FAILED.
  exit /b 1
)

echo.
echo Publish OK.
echo Output: bin\Release\net10.0-windows\win-x64\publish\
echo Executable: LMM/bin/Release/net10.0-windows/win-x64/publish/LMM.exe
endlocal
pause
