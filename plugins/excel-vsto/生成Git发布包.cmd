@echo off
setlocal
set "LOG=%~dp0github-publish-result.log"
set "PROJECT=%~dp0CellImageBridgeVsto.csproj"
set "PFX=%~dp0CellImageBridgeVsto_TemporaryKey.pfx"
set "SOURCE=%~dp0bin\Release\app.publish"
set "OUTPUT=%LOCALAPPDATA%\CellImageBridgeVsto\github-publish"
set "MSBUILD=C:\Program Files\Microsoft Visual Studio\18\Community\MSBuild\Current\Bin\MSBuild.exe"

echo Checking the development signing certificate...
certutil -user -store My F26B0873F7E4D96192FA33A7294EB68ABFECD6D6 > "%LOG%" 2>&1
if errorlevel 1 certutil -user -f -p "" -importPFX My "%PFX%" > "%LOG%" 2>&1
if errorlevel 1 goto failed

echo Publishing VSTO package with GitHub download URL...
"%MSBUILD%" "%PROJECT%" /t:Publish /p:Configuration=Release /p:VisualStudioVersion=18.0 /p:InstallUrl="https://raw.githubusercontent.com/gangchen0470/excel-wps-cell-image-bridge/v1/wps-floating-foundation/vsto/" /nologo /verbosity:minimal >> "%LOG%" 2>&1
if errorlevel 1 goto failed

echo Copying package to a short staging path...
if exist "%OUTPUT%" rmdir /s /q "%OUTPUT%"
mkdir "%OUTPUT%" >nul 2>&1
xcopy "%SOURCE%\*" "%OUTPUT%\" /e /i /y >nul
copy /y "%~dp0bin\Release\Microsoft.Office.Tools.Common.v4.0.Utilities.dll" "%OUTPUT%\Application Files\CellImageBridgeVsto_1_0_1_0\Microsoft.Office.Tools.Common.v4.0.Utilities.dll.deploy" >nul
if not exist "%OUTPUT%\CellImageBridgeVsto.vsto" goto failed
if not exist "%OUTPUT%\Application Files\CellImageBridgeVsto_1_0_1_0\Microsoft.Office.Tools.Common.v4.0.Utilities.dll.deploy" goto failed

echo GitHub VSTO package generated successfully:
echo %OUTPUT%
pause
exit /b 0

:failed
echo.
echo GitHub VSTO package generation failed:
type "%LOG%"
pause
exit /b 1
