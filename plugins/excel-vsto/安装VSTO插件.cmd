@echo off
setlocal
set "LOG=%~dp0vsto-install-result.log"
set "PROJECT=%~dp0CellImageBridgeVsto.csproj"
set "PFX=%~dp0CellImageBridgeVsto_TemporaryKey.pfx"
set "PUBLISH_SOURCE=%~dp0bin\Release\app.publish"
set "INSTALL_ROOT=%LOCALAPPDATA%\CellImageBridgeVsto\publish"
set "MANIFEST=%INSTALL_ROOT%\CellImageBridgeVsto.vsto"
set "MSBUILD=C:\Program Files\Microsoft Visual Studio\18\Community\MSBuild\Current\Bin\MSBuild.exe"
set "VSTOINSTALLER=C:\Program Files\Common Files\microsoft shared\VSTO\10.0\VSTOInstaller.exe"

echo Closing Excel processes...
taskkill /f /im excel.exe >nul 2>&1

echo Checking the local development signing certificate...
certutil -user -store My F26B0873F7E4D96192FA33A7294EB68ABFECD6D6 > "%LOG%" 2>&1
if errorlevel 1 certutil -user -f -p "" -importPFX My "%PFX%" > "%LOG%" 2>&1
if errorlevel 1 goto certificate_failed

echo Publishing signed VSTO package...
"%MSBUILD%" "%PROJECT%" /t:Publish /p:Configuration=Release /p:VisualStudioVersion=18.0 /nologo /verbosity:minimal >> "%LOG%" 2>&1
if errorlevel 1 goto build_failed

if not exist "%PUBLISH_SOURCE%\CellImageBridgeVsto.vsto" goto manifest_missing

echo Copying the package to a short local path...
if exist "%INSTALL_ROOT%" rmdir /s /q "%INSTALL_ROOT%"
mkdir "%INSTALL_ROOT%" >nul 2>&1
xcopy "%PUBLISH_SOURCE%\*" "%INSTALL_ROOT%\" /e /i /y >nul
if errorlevel 1 goto copy_failed
copy /y "%~dp0bin\Release\Microsoft.Office.Tools.Common.v4.0.Utilities.dll" "%INSTALL_ROOT%\Application Files\CellImageBridgeVsto_1_0_1_0\Microsoft.Office.Tools.Common.v4.0.Utilities.dll.deploy" >nul
if errorlevel 1 goto copy_failed
if not exist "%INSTALL_ROOT%\Application Files\CellImageBridgeVsto_1_0_1_0\Microsoft.Office.Tools.Common.v4.0.Utilities.dll.deploy" goto copy_failed
if not exist "%MANIFEST%" goto manifest_missing

echo Removing legacy COM registration...
reg delete "HKCU\Software\Microsoft\Office\Excel\Addins\CellImageBridge.ExcelAddIn" /f >nul 2>&1
reg delete "HKCU\Software\Classes\CellImageBridge.ExcelAddIn" /f >nul 2>&1
reg delete "HKCU\Software\Classes\CLSID\{BA5DB613-0DDA-4F45-9270-FA308C186342}" /f >nul 2>&1

echo Opening Microsoft VSTO Installer. Click Install in the dialog.
"%VSTOINSTALLER%" /Install "%MANIFEST%"
if errorlevel 1 goto install_failed

echo.
echo VSTO add-in installation completed.
echo Log: %LOG%
pause
exit /b 0

:certificate_failed
echo.
echo Could not import the development signing certificate:
type "%LOG%"
pause
exit /b 1

:build_failed
echo.
echo VSTO publish failed:
type "%LOG%"
pause
exit /b 1

:copy_failed
echo.
echo Could not copy the deployment package to %INSTALL_ROOT%
pause
exit /b 1

:manifest_missing
echo.
echo Deployment manifest was not generated.
type "%LOG%"
pause
exit /b 1

:install_failed
echo.
echo VSTOInstaller did not complete.
pause
exit /b 1
