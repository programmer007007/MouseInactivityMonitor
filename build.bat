@echo off
setlocal

rem Build and deploy MouseInactivityMonitor

taskkill /F /IM MouseInactivityMonitor.exe >nul 2>&1

pushd "%~dp0"

dotnet clean
if errorlevel 1 goto :error

dotnet publish -c Release -r win-x64 --self-contained true -p:PublishSingleFile=true
if errorlevel 1 goto :error

if not exist "C:\MouseWatcher" mkdir "C:\MouseWatcher"
copy /Y "%~dp0bin\Release\net8.0-windows\win-x64\publish\MouseInactivityMonitor.exe" "C:\MouseWatcher\MouseInactivityMonitor.exe"
if errorlevel 1 goto :error
copy /Y "%~dp0fan.png" "C:\MouseWatcher\fan.png"
if errorlevel 1 goto :error

start "" "C:\MouseWatcher\MouseInactivityMonitor.exe" --tray

popd
echo Done.
pause
exit /b 0

:error
popd
echo Build failed.
pause
exit /b 1
