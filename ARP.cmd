@setlocal DisableDelayedExpansion
@echo off
set auto=0
if /i "%~1"=="/auto" (
set auto=1
)

set "SysPath=%SystemRoot%\System32"
set "Path=%SystemRoot%\System32;%SystemRoot%\System32\Wbem"
if exist "%SystemRoot%\Sysnative\reg.exe" (
set "SysPath=%SystemRoot%\Sysnative"
set "Path=%SystemRoot%\Sysnative;%SystemRoot%\Sysnative\Wbem;%Path%"
)
set "_temp=%temp%"

set _xp=0
REG QUERY "HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion" /v CurrentBuildNumber |FINDSTR 2600 >NUL
if %errorlevel% equ 0 (
if %auto% equ 1 goto :eof
set _xp=1
echo ==== Notice ====
echo This script do not support Windows XP x86 {Build 2600}
echo.
echo Press any key to exit...
pause >nul
goto :eof
)
reg query "HKU\S-1-5-19" 1>nul 2>nul || (
echo ==== ERROR ====
echo This script require administrator privileges.
echo To do so, right click on this script and select 'Run as administrator'
echo.
echo Press any key to exit...
pause >nul
goto :eof
)

set "arch=x64"
if /i "%PROCESSOR_ARCHITECTURE%"=="x86" if "%PROCESSOR_ARCHITEW6432%"=="" set "arch=x86"

set "_Nul1=1>nul"
set "_Nul2=2>nul"
set "_Nul6=2^>nul"
set "_Nul3=1>nul 2>nul"
setlocal EnableDelayedExpansion

set "mvc=Microsoft Visual C++"
set "_natkey=hklm\software\microsoft\windows\currentversion\uninstall"
set "_wowkey=hklm\software\wow6432node\microsoft\windows\currentversion\uninstall"

if exist "!_temp!\msi*.txt" del /f /q "!_temp!\msi*.txt"

if %arch%==x86 goto :MsiNat

:MsiWow
for %%G in (
"%mvc% 2005 Redistributable"
"%mvc% 2008 Redistributable"
"%mvc% 2010  x86 Redistributable"
"%mvc% 2012 x86 Additional Runtime"
"%mvc% 2012 x86 Minimum Runtime"
"%mvc% 2013 x86 Additional Runtime"
"%mvc% 2013 x86 Minimum Runtime"
"%mvc% 14 x86 Additional Runtime"
"%mvc% 14 x86 Minimum Runtime"
"%mvc% 2015 x86 Additional Runtime"
"%mvc% 2015 x86 Minimum Runtime"
"%mvc% 2017 x86 Additional Runtime"
"%mvc% 2017 x86 Minimum Runtime"
"%mvc% 2019 x86 Additional Runtime"
"%mvc% 2019 x86 Minimum Runtime"
"%mvc% 2022 x86 Additional Runtime"
"%mvc% 2022 x86 Minimum Runtime"
"Microsoft Visual Studio 2010 Tools for Office Runtime"
"Microsoft Visual Basic/C++ Runtime"
) do (
reg query %_wowkey% /f %%G /s %_Nul2% | find /i "HKEY_LOCAL_MACHINE" >>"!_temp!\msi32.txt"
)

:MsiNat
for %%G in (
"%mvc% 2005 Redistributable"
"%mvc% 2008 Redistributable"
"%mvc% 2010  %arch% Redistributable"
"%mvc% 2012 %arch% Additional Runtime"
"%mvc% 2012 %arch% Minimum Runtime"
"%mvc% 2013 %arch% Additional Runtime"
"%mvc% 2013 %arch% Minimum Runtime"
"%mvc% 14 %arch% Additional Runtime"
"%mvc% 14 %arch% Minimum Runtime"
"%mvc% 2015 %arch% Additional Runtime"
"%mvc% 2015 %arch% Minimum Runtime"
"%mvc% 2017 %arch% Additional Runtime"
"%mvc% 2017 %arch% Minimum Runtime"
"%mvc% 2019 %arch% Additional Runtime"
"%mvc% 2019 %arch% Minimum Runtime"
"%mvc% 2022 %arch% Additional Runtime"
"%mvc% 2022 %arch% Minimum Runtime"
"Microsoft Visual Studio 2010 Tools for Office Runtime"
"Microsoft Visual Basic/C++ Runtime"
) do (
reg query %_natkey% /f %%G /s %_Nul2% | find /i "HKEY_LOCAL_MACHINE" >>"!_temp!\msi96.txt"
)

:menu
if %auto% equ 1 goto :hide
@cls
echo.
echo Visual C++ Runtimes entries in Add/Remove Programs:
echo.
echo 1. Hide
echo 2. Show
echo.
echo ----------------------------------------------------------
choice /c 120 /n /m "Choose a menu option, or press 0 to quit: "
if errorlevel 3 goto :quit
if errorlevel 2 goto :show
if errorlevel 1 goto :hide

:hide
@cls
findstr /i "HKEY_LOCAL_MACHINE" "!_temp!\msi32.txt" %_Nul3% || goto :hidemsi
for /f "usebackq tokens=8 delims=\" %%G in ("!_temp!\msi32.txt") do (
reg add %_wowkey%\%%G /f /v SystemComponent /t REG_DWORD /d 1 %_Nul3%
)

:hidemsi
findstr /i "HKEY_LOCAL_MACHINE" "!_temp!\msi96.txt" %_Nul3% || goto :close
for /f "usebackq tokens=7 delims=\" %%G in ("!_temp!\msi96.txt") do (
reg add %_natkey%\%%G /f /v SystemComponent /t REG_DWORD /d 1 %_Nul3%
)
goto :close

:show
@cls
findstr /i "HKEY_LOCAL_MACHINE" "!_temp!\msi32.txt" %_Nul3% || goto :showmsi
for /f "usebackq tokens=8 delims=\" %%G in ("!_temp!\msi32.txt") do (
reg delete %_wowkey%\%%G /f /v SystemComponent %_Nul3%
)

:showmsi
findstr /i "HKEY_LOCAL_MACHINE" "!_temp!\msi96.txt" %_Nul3% || goto :close
for /f "usebackq tokens=7 delims=\" %%G in ("!_temp!\msi96.txt") do (
reg delete %_natkey%\%%G /f /v SystemComponent %_Nul3%
)
goto :close

:quit
if exist "!_temp!\msi*.txt" del /f /q "!_temp!\msi*.txt"
goto :eof

:close
if exist "!_temp!\msi*.txt" del /f /q "!_temp!\msi*.txt"
if %auto% equ 1 goto :eof
echo ==== Done ====
echo.
echo Press any key to exit...
pause >nul
goto :eof
