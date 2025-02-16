@echo off

SETLOCAL ENABLEDELAYEDEXPANSION
SET "currentDir=%~dp0"
SET "FolderName=Office 365 - Office Deployment Tool"
SET "FolderPath=%currentDir%%FolderName%"

:: Running as Admin privileges
NET SESSION >nul 2>&1 
IF %ERRORLEVEL% NEQ 0 (
	echo Requesting admin privileges...
	powershell -Command "Start-Process '%~0' -Verb RunAs" 
	exit /b
)

echo Running with admin privileges...
echo The script is executing with administrative rights, allowing necessary installations and deletions.
echo.
if exist "!FolderPath!" (
	echo Emptying the existing folder: %FolderPath%...
	powershell -Command "Remove-Item -Path '!FolderPath!' -Recurse -Force" 
	if %ERRORLEVEL% NEQ 0 (
		echo Failed to empty the folder.
		pause
		exit /b
	)
	echo.
)

echo Choose the installation channel for Office 365:
echo [1] Stable Release
echo [2] Beta Release
echo.
echo If you are unsure, choose 1.
echo.

:: Use the choice command for single-key input
choice /C 12 /N /M "Option (1 or 2):"
SET "choice=%errorlevel%"

SET "setupURL=https://officecdn.microsoft.com/pr/wsus/setup.exe"

if "%choice%"=="1" (
	SET "channelType=CurrentPreview"
) else if "%choice%"=="2" (
	SET "channelType=BetaChannel"
) else (
	echo.
	echo Invalid choice! Exiting...
	pause
	exit /b
)
echo You selected the %channelType% version for installation.
mkdir "!FolderPath!"
echo Folder: "%FolderName%" created at "%currentDir%".

:: Download the setup.exe
echo Downloading "setup.exe"...
powershell -Command "Invoke-WebRequest -Uri '%setupURL%' -OutFile '!FolderPath!\setup.exe'"
if %ERRORLEVEL% NEQ 0 (
	echo.
	echo Failed to download setup.exe. Exiting...
	pause
	exit /b
)
echo.
echo The setup.exe and XML configuration - ("%channelType%") has been successfully downloaded/created at:
echo "%FolderPath%" and will be use for storing installation files

:: Create the XML configuration
(
echo <Configuration ID="2526aef5-bb32-48b8-8b28-11dd9740b2df">
echo	^<Remove All="True"^>
echo		^<RemoveMSI All="True" /^>
echo	^</Remove^>
echo	<Info Description="This Office365 was installed with custom configuration." />
echo	<Add OfficeClientEdition="64" Channel="!channelType!" MigrateArch="TRUE">
echo		<Product ID="O365ProPlusRetail">
echo		<Language ID="MatchOS" />
echo		<Language ID="MatchPreviousMSI" />
echo		<ExcludeApp ID="Access" />
echo		<ExcludeApp ID="Groove" />
echo		<ExcludeApp ID="Lync" />
echo		<ExcludeApp ID="OneDrive" />
echo		<ExcludeApp ID="OneNote" />
echo		<ExcludeApp ID="Outlook" />
echo		<ExcludeApp ID="Publisher" />
echo		<ExcludeApp ID="Bing" />
echo	</Product>
echo	</Add>
echo	<Property Name="SharedComputerLicensing" Value="0" />
echo	<Property Name="FORCEAPPSHUTDOWN" Value="TRUE" />
echo	<Property Name="DeviceBasedLicensing" Value="0" />
echo	<Property Name="SCLCacheOverride" Value="0" />
echo	<Updates Enabled="TRUE" />
echo	<RemoveMSI />
echo	<AppSettings>
echo		<Setup Name="Company" Value="MSO-365" />
echo	</AppSettings>
echo	<Display Level="Full" AcceptEULA="TRUE" />
echo </Configuration>
) > "!FolderPath!\!channelType!.xml"

cd /d "!FolderPath!" || (
	echo.
	echo Failed to navigate to the directory. Exiting...
	pause
	exit /b
)

echo.
echo Initiating the download of essential files for the Office 365 installation.
setup.exe /download "!channelType!.xml"
echo All necessary files have been downloaded successfully.

echo.
echo Configuring Office 365 installation...
setup.exe /configure "!channelType!.xml"
echo The configuration is complete, and Office 365 is ready for use.

echo.
echo The installation was successful. You can start using Office 365 right away!
echo.
echo All installation files have been removed from folder "%FolderName%" to clean up your directory.
echo.
pause
start cmd /k "cd /d !currentDir! && rmdir /s /q "%FolderPath%" && exit"
