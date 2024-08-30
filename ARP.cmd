@setlocal DisableDelayedExpansion
@echo off
set auto=0
set bt32=0
if /i "%~1"=="/auto" (
set auto=1
)
if /i "%~2"=="/x86" set bt32=1

set "SysPath=%SystemRoot%\System32"
set "Path=%SystemRoot%\System32;%SystemRoot%;%SystemRoot%\System32\Wbem;%SystemRoot%\System32\WindowsPowerShell\v1.0\"
if exist "%SystemRoot%\Sysnative\reg.exe" (
set "SysPath=%SystemRoot%\Sysnative"
set "Path=%SystemRoot%\Sysnative;%SystemRoot%;%SystemRoot%\Sysnative\Wbem;%SystemRoot%\Sysnative\WindowsPowerShell\v1.0\;%Path%"
)
set "_temp=%temp%"

set _xp=0
REG QUERY "HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion" /v CurrentBuildNumber |FINDSTR 2600 >NUL
if %errorlevel% equ 0 set _xp=1
if %_xp% equ 1 goto :unsupported

reg.exe query "HKU\S-1-5-19" 1>nul 2>nul || goto :unadmin

set "arch=x64"
if /i "%PROCESSOR_ARCHITECTURE%"=="x86" if "%PROCESSOR_ARCHITEW6432%"=="" set "arch=x86"
if /i "%PROCESSOR_ARCHITECTURE%"=="arm64" set "arch=arm64"
if /i "%PROCESSOR_ARCHITEW6432%"=="arm64" set "arch=arm64"
if /i "%PROCESSOR_ARCHITEW6432%"=="amd64" set "arch=x64"

set "_Nul1=1>nul"
set "_Nul2=2>nul"
set "_Nul6=2^>nul"
set "_Nul3=1>nul 2>nul"
setlocal EnableDelayedExpansion

set "mvc=Microsoft Visual C++"
set "mvt=Microsoft Visual Studio 2010 Tools for Office Runtime"
set "_g_="{.*-.*-.*-.*-.*}""
set "_h_="HKEY_LOCAL_MACHINE""
set "_r_=Redistributable"
set "_l_=Additional Runtime"
set "_m_=Minimum Runtime"
set "_natkey=HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall"
set "_wowkey=HKLM\SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall"

:MsiWow
if exist "!_temp!\msi*.txt" del /f /q "!_temp!\msi*.txt"

if %arch%==x86 goto :MsiNat

call :chMSI wow

:MsiNat
set skpNat=0
if %arch%==arm64 set skpNat=1
if not %arch%==x86 if %bt32% equ 1 set skpNat=1

if %skpNat% equ 1 goto :menu

call :chMSI nat

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
findstr /i %_h_% "!_temp!\msi32.txt" %_Nul3% && for /f "usebackq tokens=8 delims=\" %%G in ("!_temp!\msi32.txt") do (
reg add %_wowkey%\%%G /f /v SystemComponent /t REG_DWORD /d 1 %_Nul3%
)
if %skpNat% equ 1 goto :close
findstr /i %_h_% "!_temp!\msi96.txt" %_Nul3% && for /f "usebackq tokens=7 delims=\" %%G in ("!_temp!\msi96.txt") do (
reg add %_natkey%\%%G /f /v SystemComponent /t REG_DWORD /d 1 %_Nul3%
)
goto :close

:show
@cls
findstr /i %_h_% "!_temp!\msi32.txt" %_Nul3% && for /f "usebackq tokens=8 delims=\" %%G in ("!_temp!\msi32.txt") do (
reg delete %_wowkey%\%%G /f /v SystemComponent %_Nul3%
)
if %skpNat% equ 1 goto :close
findstr /i %_h_% "!_temp!\msi96.txt" %_Nul3% && for /f "usebackq tokens=7 delims=\" %%G in ("!_temp!\msi96.txt") do (
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
goto :E_Exit

:unadmin
echo ==== ERROR ====
echo This script require administrator privileges.
echo To do so, right click on this script and select 'Run as administrator'
goto :E_Exit

:unsupported
if %auto% equ 1 goto :eof
echo ==== Notice ====
echo This script do not support Windows XP x86 {Build 2600}
goto :E_Exit

:E_Exit
echo.
echo Press any key to exit.
pause >nul
exit /b

:chMSI
if "%1"=="nat" (
set _a_=%arch%
set _k_=%_natkey%
set _f_=msi96
) else (
set _a_=x86
set _k_=%_wowkey%
set _f_=msi32
)

for %%G in (
"%mvc% 2005 %_r_%"
"%mvc% 2008 %_r_%"
"%mvc% 2010  %_a_% %_r_%"
"%mvc% 2012 %_a_% %_l_%"
"%mvc% 2012 %_a_% %_m_%"
"%mvc% 2013 %_a_% %_l_%"
"%mvc% 2013 %_a_% %_m_%"
"%mvc% 14 %_a_% %_l_%"
"%mvc% 14 %_a_% %_m_%"
"%mvc% 2015 %_a_% %_l_%"
"%mvc% 2015 %_a_% %_m_%"
"%mvc% 2017 %_a_% %_l_%"
"%mvc% 2017 %_a_% %_m_%"
"%mvc% 2019 %_a_% %_l_%"
"%mvc% 2019 %_a_% %_m_%"
"%mvc% 2022 %_a_% %_l_%"
"%mvc% 2022 %_a_% %_m_%"
"%mvt%"
"Microsoft Visual Basic/C++ Runtime"
"Microsoft Visual Basic Runtime"
"Microsoft Visual C++ 2002-2003 Runtime"
) do (
reg query %_k_% /f %%G /s %_Nul2% | find /i %_h_% | findstr /r %_g_% >>"!_temp!\%_f_%.txt"
)

goto :eof
