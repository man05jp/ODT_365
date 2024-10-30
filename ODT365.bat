@echo off
SETLOCAL ENABLEDELAYEDEXPANSION

:: Set the raw GitHub URL
SET "fileURL=https://raw.githubusercontent.com/Mano5JP/ODT_365/refs/heads/main/ODT"
SET "FileNameOf=OfficeDeploymentTool.bat"
:: Download the latest batch file
echo Downloading the latest version of the script...
powershell -Command "Invoke-WebRequest -Uri '!fileURL!' -OutFile %FileNameOf%"

if %ERRORLEVEL% NEQ 0 (
    echo Failed to download the latest version.
    pause
    exit /b
)

echo Download complete. Running the latest version...
call "%FileNameOf%"
