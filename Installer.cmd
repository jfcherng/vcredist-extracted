@setlocal DisableDelayedExpansion
@echo off
set _debug=0
set vci=v0.53.0
set auto=0
set verbosity=/quiet
set verbosityshort=/qn /norestart
if /i "%~1"=="/auto" (
set auto=1
set verbosity=/passive
set verbosityshort=/qb
)
if /i "%~1"=="/ucrt" (
set auto=1
set ucrt=1
set verbosity=/passive
set verbosityshort=/qb
)
if /i "%~1"=="/quiet" (
set auto=1
)
if /i "%~1"=="/vcpp" (
set auto=1
set vcpp=1
)
if /i "%~1"=="/uc14" (
set auto=1
set uc14=1
)
if /i "%~1"=="/debug" (
set _debug=1
)
set installcount=0
set count=0
set invalid=0

set "SysPath=%SystemRoot%\System32"
if exist "%SystemRoot%\Sysnative\reg.exe" (set "SysPath=%SystemRoot%\Sysnative")
set "Path=%SysPath%;%SystemRoot%;%SysPath%\Wbem"
set "_temp=%temp%"
set "_work=%~dp0"
set "_work=%_work:~0,-1%"
for /f "tokens=2*" %%a in ('reg query "HKCU\Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders" /v Desktop') do call set "_log=%%b\VCpp"

set "arch=x64"&set "xBT=amd64"
if /i "%PROCESSOR_ARCHITECTURE%"=="x86" if "%PROCESSOR_ARCHITEW6432%"=="" set "arch=x86"&set "xBT=x86"
if /i "%PROCESSOR_ARCHITECTURE%"=="arm64" set "arch=arm64"&set "xBT=x86"
if /i "%PROCESSOR_ARCHITEW6432%"=="arm64" set "arch=arm64"&set "xBT=x86"
set wixpkg=vcredist_x86.exe,vcredist_x64.exe,vc_redist.x86.exe,vc_redist.x64.exe
if %arch%==arm64 set wixpkg=vcredist_x86.exe,vc_redist.x86.exe

set _xp=0
ver|findstr /c:" 5." >nul
if %errorlevel% equ 0 (
if %auto% equ 1 goto :eof
echo ==== Notice ====
echo VisualCppRedist_AIO v35 is the last version to support Windows XP
echo VC++ 2019 v14.28.29213.0 is the last version compatible with Windows XP
echo.
if %_debug% equ 1 goto :eof
echo Press any key to exit...
pause >nul
goto :eof
)

if %_debug% equ 0 reg query "HKU\S-1-5-19" 1>nul 2>nul || (
echo ==== ERROR ====
echo This script require administrator privileges.
echo To do so, right click on this script and select 'Run as administrator'
echo.
echo Press any key to exit...
pause >nul
goto :eof
)

setlocal EnableDelayedExpansion
if %_debug% equ 0 (
  set "_Nul1=1>nul"
  set "_Nul2=2>nul"
  set "_Nul6=2^>nul"
  set "_Nul3=1>nul 2>nul"
  goto :Begin
)
  set "_Nul1="
  set "_Nul2="
  set "_Nul6="
  set "_Nul3="
  set auto=1
echo.
echo Running in Debug Mode...
echo The window will be closed when finished
@echo on
@prompt $G
@call :Begin >"!_log!_tmp.log" 2>&1 &cmd /u /c type "!_log!_tmp.log">"!_log!_Debug.log"&del "!_log!_tmp.log"
@title %ComSpec%
@echo off
@exit /b

:Begin
if defined uc14 goto :ucrtonly
title Visual C++ Redistributable AIO %vci%

for /f "tokens=6 delims=[]. " %%G in ('ver') do set winbuild=%%G

for /f "skip=2 tokens=3* delims= " %%G in ('"reg query "hklm\software\microsoft\Windows NT\currentversion" /v productname" %_Nul6%') do set "winv=%%G %%H"

if %winbuild% geq 7601 for /f "tokens=3" %%G in ('"reg query "hklm\software\microsoft\Windows NT\currentversion" /v UBR" %_Nul6%') do if not errorlevel 1 set /a "UBR=%%G"

if defined UBR (
for /f "skip=2 tokens=3,4,6,7 delims=. " %%G in ('reg query "hklm\software\microsoft\Windows NT\currentversion" /v buildlabex') do set "_os=%winv% %arch% {%%G.%UBR%.%%J.%%I}"
) else (
for /f "skip=2 tokens=3,4,6,7 delims=. " %%G in ('reg query "hklm\software\microsoft\Windows NT\currentversion" /v buildlabex') do set "_os=%winv% %arch% {%%G.%%H.%%J.%%I}"
)

if %auto% equ 1 goto :proceed

:top
if %_debug% equ 1 goto :proceed
set _inp=
call :title
echo Before installing the latest Visual C++ Runtimes
echo any existing non-compliant versions will be removed.
echo.
echo These include the original EXE installers or the MSI packages
echo with a version lower than the ones bundled with this Installer.
echo.
echo Proceed with the check/removal and installation?
echo.
echo ----------------------------------------------------------
choice /c YRN /n /m "Press Y for Yes, R for ReadMe, or N to cancel and exit> "
if errorlevel 3 goto :eof
if errorlevel 2 goto :readme
if errorlevel 1 goto :proceed
goto :top

:proceed
pushd "!_work!"
set "mvc=Microsoft Visual C++"

set "_val=/v UninstallString"
set "_natkey=hklm\software\microsoft\windows\currentversion\uninstall"
set "_wowkey=hklm\software\wow6432node\microsoft\windows\currentversion\uninstall"

set "_vstor=608280"
set "_ver08=507276229"
set "_ver09=307297523"
set "_ver10=40219473"
set "_ver11=61135400"
set "_ver12=406640"
set "_ver14=30305280"

set "_filevstor=%CommonProgramFiles%\Microsoft Shared\VSTO\vstoee.dll"

set "_x86file08=%SystemRoot%\WinSxS\x86_microsoft.vc80.crt_1fc8b3b9a1e18e3b_8.0.50727.6229_none_d089f796442de10e\msvcp80.dll"
set "_x86file09=%SystemRoot%\WinSxS\x86_microsoft.vc90.crt_1fc8b3b9a1e18e3b_9.0.30729.7523_none_508f21ccbcbbb7a8\msvcp90.dll"
set "_x86file10=msvcp100.dll"
set "_x86file11=msvcp110.dll"
set "_x86file12=msvcp120.dll"
set "_x86file14=msvcp140.dll"

set "_x64file08=%SystemRoot%\WinSxS\amd64_microsoft.vc80.crt_1fc8b3b9a1e18e3b_8.0.50727.6229_none_88dcc0bf2fb1b808\msvcp80.dll"
set "_x64file09=%SystemRoot%\WinSxS\amd64_microsoft.vc90.crt_1fc8b3b9a1e18e3b_9.0.30729.7523_none_08e1eaf5a83f8ea2\msvcp90.dll"
set "_x64file10=msvcp100.dll"
set "_x64file11=msvcp110.dll"
set "_x64file12=msvcp120.dll"
set "_x64file14=msvcp140.dll"

set "_x86code08={710f4c1c-cc18-4c49-8cbf-51240c89a1a2}"
set "_x86code09={9BE518E6-ECC6-35A9-88E4-87755C07200F}"
set "_x86code10={F0C3E5D1-1ADE-321E-8167-68EF0DE699A5}"
set "_x86code10c=1D5E3C0FEDA1E123187686FED06E995A"
set "_x86code11m={BD95A8CD-1D9F-35AD-981A-3E7925026EBB}"
set "_x86code11a={B175520C-86A2-35A7-8619-86DC379688B9}"
set "_x86code12m={8122DAB1-ED4D-3676-BB0A-CA368196543E}"
set "_x86code12a={D401961D-3A20-3AC7-943B-6139D5BD490A}"
set "_x86code14m={76EE50C2-B749-49FF-A512-1DFE6D67AA1C}"
set "_x86code14a={32571C34-D349-4598-A97F-D621A3070FAB}"

set "_x64code08={ad8a2fa1-06e7-4b0d-927d-6e54b3d31028}"
set "_x64code09={5FCE6D76-F5DC-37AB-B2B8-22AB8CEDB1D4}"
set "_x64code10={1D8E6291-B0D5-35EC-8441-6616F567A0F7}"
set "_x64code10c=1926E8D15D0BCE53481466615F760A7F"
set "_x64code11m={CF2BEA3C-26EA-32F8-AA9B-331F7E34BA97}"
set "_x64code11a={37B8F9C7-03FB-3253-8781-2517C99D7C00}"
set "_x64code12m={53CF6934-A98D-3D84-9146-FC4EDF3D5641}"
set "_x64code12a={010792BA-551A-3AC0-A7EF-0FAB4156C382}"
set "_x64code14m={91823FBE-6862-491D-A419-E2E9D868AEB3}"
set "_x64code14a={2470F04D-028C-49C6-BF05-6D27A2C9C348}"

if exist "!_temp!\msi.txt" del /f /q "!_temp!\msi.txt"
if exist "!_temp!\wix.txt" del /f /q "!_temp!\wix.txt"
if exist filever.vbs del /f /q filever.vbs

echo>>filever.vbs Set objFSO = CreateObject^("Scripting.FileSystemObject"^)
echo>>filever.vbs Wscript.Echo objFSO.GetFileVersion^(WScript.arguments^(0^)^)

set "RegKey=SOFTWARE\Microsoft\Windows Script Host\Settings"
reg query "HKCU\%RegKey%" /v Enabled %_Nul2% | find /i "0x0" %_Nul1% && (set vbscu=1&reg delete "HKCU\%RegKey%" /v Enabled /f %_Nul3%)
reg query "HKLM\%RegKey%" /v Enabled %_Nul2% | find /i "0x0" %_Nul1% && (set vbslm=1&reg delete "HKLM\%RegKey%" /v Enabled /f %_Nul3%)

if %arch%==x86 goto :WiXNat

:WiX32
call :title

for %%G in (
"%mvc% 2012 Redistributable"
"%mvc% 2013 Preview Redistributable"
"%mvc% 2013 RC Redistributable"
"%mvc% 2013 Redistributable"
"%mvc% 14 CTP Redistributable"
"%mvc% 2015 Preview Redistributable"
"%mvc% 2015 CTP Redistributable"
"%mvc% 2015 RC Redistributable"
"%mvc% 2015 Redistributable"
"%mvc% 2017 RC Redistributable"
"%mvc% 2017 Redistributable"
"%mvc% 2019 Redistributable"
"%mvc% 2022 Redistributable"
"%mvc% 2015-2019 Redistributable"
"%mvc% 2015-2022 Redistributable"
) do (
reg query %_wowkey% /f %%G /s %_Nul2% | find /i "HKEY_LOCAL_MACHINE" >>"!_temp!\wix.txt"
)

findstr /i "HKEY_LOCAL_MACHINE" "!_temp!\wix.txt" %_Nul3% || goto :Msi32

echo Uninstalling non-compliant Visual C++ WiX packages {x64/x86}
echo ^(please wait as this process may take a few moments^)

set invalid=1
for /f "usebackq tokens=8 delims=\" %%G in ("!_temp!\wix.txt") do (
for %%H in (%wixpkg%) do (
  if exist "%ProgramData%\Package Cache\%%G\%%H" (
    if %_debug% equ 0 (
      start /wait "" "%ProgramData%\Package Cache\%%G\%%H" /uninstall %verbosity% /norestart
      reg delete %_wowkey%\%%G /f %_Nul3%
      )
    )
  )
)

:Msi32
call :title

for %%G in (08,09,10,11,12,14) do set _x86install%%G=1

if exist "%SystemRoot%\SysWOW64\!_x86file10!" for /f "tokens=3,4 delims=." %%i in ('cscript //nologo filever.vbs "%SystemRoot%\SysWOW64\!_x86file10!"') do (
if %%i gtr %_ver10:~0,5% set _x86install10=0
if %%i equ %_ver10:~0,5% if %%j geq %_ver10:~5,3% set _x86install10=0
)
if exist "%SystemRoot%\SysWOW64\!_x86file11!" for /f "tokens=3,4 delims=." %%i in ('cscript //nologo filever.vbs "%SystemRoot%\SysWOW64\!_x86file11!"') do (
if %%i gtr %_ver11:~0,5% set _x86install11=0
if %%i equ %_ver11:~0,5% if %%j geq %_ver11:~5,3% set _x86install11=0
)
if exist "%SystemRoot%\SysWOW64\!_x86file12!" for /f "tokens=3,4 delims=." %%i in ('cscript //nologo filever.vbs "%SystemRoot%\SysWOW64\!_x86file12!"') do (
if %%i gtr %_ver12:~0,5% set _x86install12=0
if %%i equ %_ver12:~0,5% if %%j geq %_ver12:~5,1% set _x86install12=0
)
if exist "%SystemRoot%\SysWOW64\!_x86file14!" for /f "tokens=2-4 delims=." %%i in ('cscript //nologo filever.vbs "%SystemRoot%\SysWOW64\!_x86file14!"') do (
if %%i gtr %_ver14:~0,2% set _x86install14=0
if %%i equ %_ver14:~0,2% if %%j gtr %_ver14:~2,5% set _x86install14=0
if %%i equ %_ver14:~0,2% if %%j equ %_ver14:~2,5% if %%k geq %_ver14:~7,1% set _x86install14=0
)
for %%G in (08,09) do (
if exist "!_x86file%%G!" set _x86install%%G=0
)
for %%G in (08,09) do if !_x86install%%G! equ 0 (
reg query %_wowkey%\!_x86code%%G! %_val% %_Nul3% || set _x86install%%G=1
)
if !_x86install10! equ 0 (
reg query %_wowkey%\%_x86code10% %_val% %_Nul3% || set _x86install10=1
reg query HKLM\SOFTWARE\Classes\Installer\Features\%_x86code10c% /v "VC_RED_enu_x86_net_SETUP" %_Nul3% && set _x86install10=1
)
for %%G in (11,12,14) do if !_x86install%%G! equ 0 (
reg query %_wowkey%\!_x86code%%Gm! %_val% %_Nul3% || set _x86install%%G=1
reg query %_wowkey%\!_x86code%%Ga! %_val% %_Nul3% || set _x86install%%G=1
)

reg query %_wowkey% /f "%mvc% 2005 Redistributable" /s %_Nul2% | find /i "HKEY_LOCAL_MACHINE" | findstr /i /v %_x86code08% >>"!_temp!\msi.txt"

reg query %_wowkey% /f "%mvc% 2008 Redistributable" /s %_Nul2% | find /i "HKEY_LOCAL_MACHINE" | findstr /i /v %_x86code09% >>"!_temp!\msi.txt"

reg query %_wowkey% /f "%mvc% 2010  x86 Redistributable" /s %_Nul2% | find /i "HKEY_LOCAL_MACHINE" | findstr /i /v %_x86code10% >>"!_temp!\msi.txt"
if %_x86install10% equ 1 reg query %_wowkey% /f "%mvc% 2010  x86 Redistributable" /s %_Nul2% | find /i "HKEY_LOCAL_MACHINE" >>"!_temp!\msi.txt"

reg query %_wowkey% /f "%mvc% 2012 x86 Additional Runtime" /s %_Nul2% | find /i "HKEY_LOCAL_MACHINE" | findstr /i /v %_x86code11a% >>"!_temp!\msi.txt"
reg query %_wowkey% /f "%mvc% 2012 x86 Minimum Runtime" /s %_Nul2% | find /i "HKEY_LOCAL_MACHINE" | findstr /i /v %_x86code11m% >>"!_temp!\msi.txt"

reg query %_wowkey% /f "%mvc% 2013 x86 Additional Runtime" /s %_Nul2% | find /i "HKEY_LOCAL_MACHINE" | findstr /i /v %_x86code12a% >>"!_temp!\msi.txt"
reg query %_wowkey% /f "%mvc% 2013 x86 Minimum Runtime" /s %_Nul2% | find /i "HKEY_LOCAL_MACHINE" | findstr /i /v %_x86code12m% >>"!_temp!\msi.txt"

reg query %_wowkey% /f "%mvc% 2022 x86 Additional Runtime" /s %_Nul2% | find /i "HKEY_LOCAL_MACHINE" | findstr /i /v %_x86code14a% >>"!_temp!\msi.txt"
reg query %_wowkey% /f "%mvc% 2022 x86 Minimum Runtime" /s %_Nul2% | find /i "HKEY_LOCAL_MACHINE" | findstr /i /v %_x86code14m% >>"!_temp!\msi.txt"
for %%G in (
"%mvc% 14 x86 Additional Runtime"
"%mvc% 14 x86 Minimum Runtime"
"%mvc% 2015 x86 Additional Runtime"
"%mvc% 2015 x86 Minimum Runtime"
"%mvc% 2017 x86 Additional Runtime"
"%mvc% 2017 x86 Minimum Runtime"
"%mvc% 2019 x86 Additional Runtime"
"%mvc% 2019 x86 Minimum Runtime"
) do (
reg query %_wowkey% /f %%G /s %_Nul2% | find /i "HKEY_LOCAL_MACHINE" >>"!_temp!\msi.txt"
)

findstr /i "HKEY_LOCAL_MACHINE" "!_temp!\msi.txt" %_Nul3% || goto :WiXNat

echo Uninstalling non-compliant Visual C++ MSI packages {x86}
echo ^(please wait as this process may take a few moments^)

set invalid=1
for /f "usebackq tokens=8 delims=\" %%G in ("!_temp!\msi.txt") do (
if %_debug% equ 0 (
  start /wait msiexec /X%%G %verbosity% /norestart
  reg delete %_wowkey%\%%G /f %_Nul3%
  )
)

:WiXNat
call :title

if exist "!_temp!\msi.txt" del /f /q "!_temp!\msi.txt"
if exist "!_temp!\wix.txt" del /f /q "!_temp!\wix.txt"

if not %arch%==x86 goto :MsiNat

for %%G in (
"%mvc% 2012 Redistributable"
"%mvc% 2013 Preview Redistributable"
"%mvc% 2013 RC Redistributable"
"%mvc% 2013 Redistributable"
"%mvc% 14 CTP Redistributable"
"%mvc% 2015 Preview Redistributable"
"%mvc% 2015 CTP Redistributable"
"%mvc% 2015 RC Redistributable"
"%mvc% 2015 Redistributable"
"%mvc% 2017 RC Redistributable"
"%mvc% 2017 Redistributable"
"%mvc% 2019 Redistributable"
"%mvc% 2022 Redistributable"
"%mvc% 2015-2019 Redistributable"
"%mvc% 2015-2022 Redistributable"
) do (
reg query %_natkey% /f %%G /s %_Nul2% | find /i "HKEY_LOCAL_MACHINE" >>"!_temp!\wix.txt"
)
findstr /i "HKEY_LOCAL_MACHINE" "!_temp!\wix.txt" %_Nul3% || goto :MsiNat

echo Uninstalling non-compliant Visual C++ WiX packages {x86}
echo ^(please wait as this process may take a few moments^)

set invalid=1
for /f "usebackq tokens=7 delims=\" %%G in ("!_temp!\wix.txt") do (
for %%H in (vcredist_x86.exe,vc_redist.x86.exe) do (
  if exist "%ProgramData%\Package Cache\%%G\%%H" (
    if %_debug% equ 0 (
      start /wait "" "%ProgramData%\Package Cache\%%G\%%H" /uninstall %verbosity% /norestart
      reg delete %_natkey%\%%G /f %_Nul3%
      )
    )
  )
)

:MsiNat
call :title

if %arch%==arm64 (
for %%G in (08,09,10,11,12,14,vstor) do set _%arch%install%%G=0
goto :process
)

for %%G in (08,09,10,11,12,14,vstor) do set _%arch%install%%G=1

if exist "%_filevstor%" for /f "tokens=3,4 delims=." %%i in ('cscript //nologo filever.vbs "%_filevstor%"') do if %%i%%j geq %_vstor% set _%arch%installvstor=0

if exist "%SysPath%\!_%arch%file10!" for /f "tokens=3,4 delims=." %%i in ('cscript //nologo filever.vbs "%SystemRoot%\System32\!_%arch%file10!"') do (
if %%i gtr %_ver10:~0,5% set _%arch%install10=0
if %%i equ %_ver10:~0,5% if %%j geq %_ver10:~5,3% set _%arch%install10=0
)
if exist "%SysPath%\!_%arch%file11!" for /f "tokens=3,4 delims=." %%i in ('cscript //nologo filever.vbs "%SystemRoot%\System32\!_%arch%file11!"') do (
if %%i gtr %_ver11:~0,5% set _%arch%install11=0
if %%i equ %_ver11:~0,5% if %%j geq %_ver11:~5,3% set _%arch%install11=0
)
if exist "%SysPath%\!_%arch%file12!" for /f "tokens=3,4 delims=." %%i in ('cscript //nologo filever.vbs "%SystemRoot%\System32\!_%arch%file12!"') do (
if %%i gtr %_ver12:~0,5% set _%arch%install12=0
if %%i equ %_ver12:~0,5% if %%j geq %_ver12:~5,1% set _%arch%install12=0
)
if exist "%SysPath%\!_%arch%file14!" for /f "tokens=2-4 delims=." %%i in ('cscript //nologo filever.vbs "%SystemRoot%\System32\!_%arch%file14!"') do (
if %%i gtr %_ver14:~0,2% set _%arch%install14=0
if %%i equ %_ver14:~0,2% if %%j gtr %_ver14:~2,5% set _%arch%install14=0
if %%i equ %_ver14:~0,2% if %%j equ %_ver14:~2,5% if %%k geq %_ver14:~7,1% set _%arch%install14=0
)
for %%G in (08,09) do (
if exist "!_%arch%file%%G!" set _%arch%install%%G=0
)
for %%G in (08,09) do if !_%arch%install%%G! equ 0 (
reg query %_natkey%\!_%arch%code%%G! %_val% %_Nul3% || set _%arch%install%%G=1
)
if !_%arch%install10! equ 0 (
reg query %_natkey%\!_%arch%code10! %_val% %_Nul3% || set _%arch%install10=1
reg query HKLM\SOFTWARE\Classes\Installer\Features\!_%arch%code10c! /v "VC_RED_enu_%xBT%_net_SETUP" %_Nul3% && set _%arch%install10=1
)
for %%G in (11,12,14) do if !_%arch%install%%G! equ 0 (
reg query %_natkey%\!_%arch%code%%Gm! %_val% %_Nul3% || set _%arch%install%%G=1
reg query %_natkey%\!_%arch%code%%Ga! %_val% %_Nul3% || set _%arch%install%%G=1
)

reg query %_natkey% /f "%mvc% 2005 Redistributable" /s %_Nul2% | find /i "HKEY_LOCAL_MACHINE" | findstr /i /v !_%arch%code08! >>"!_temp!\msi.txt"

reg query %_natkey% /f "%mvc% 2008 Redistributable" /s %_Nul2% | find /i "HKEY_LOCAL_MACHINE" | findstr /i /v !_%arch%code09! >>"!_temp!\msi.txt"

reg query %_natkey% /f "%mvc% 2010  %arch% Redistributable" /s %_Nul2% | find /i "HKEY_LOCAL_MACHINE" | findstr /i /v !_%arch%code10! >>"!_temp!\msi.txt"
if !_%arch%install10! equ 1 reg query %_natkey% /f "%mvc% 2010  %arch% Redistributable" /s %_Nul2% | find /i "HKEY_LOCAL_MACHINE" >>"!_temp!\msi.txt"

reg query %_natkey% /f "%mvc% 2012 %arch% Additional Runtime" /s %_Nul2% | find /i "HKEY_LOCAL_MACHINE" | findstr /i /v !_%arch%code11a! >>"!_temp!\msi.txt"
reg query %_natkey% /f "%mvc% 2012 %arch% Minimum Runtime" /s %_Nul2% | find /i "HKEY_LOCAL_MACHINE" | findstr /i /v !_%arch%code11m! >>"!_temp!\msi.txt"

reg query %_natkey% /f "%mvc% 2013 %arch% Additional Runtime" /s %_Nul2% | find /i "HKEY_LOCAL_MACHINE" | findstr /i /v !_%arch%code12a! >>"!_temp!\msi.txt"
reg query %_natkey% /f "%mvc% 2013 %arch% Minimum Runtime" /s %_Nul2% | find /i "HKEY_LOCAL_MACHINE" | findstr /i /v !_%arch%code12m! >>"!_temp!\msi.txt"

reg query %_natkey% /f "%mvc% 2022 %arch% Additional Runtime" /s %_Nul2% | find /i "HKEY_LOCAL_MACHINE" | findstr /i /v !_%arch%code14a! >>"!_temp!\msi.txt"
reg query %_natkey% /f "%mvc% 2022 %arch% Minimum Runtime" /s %_Nul2% | find /i "HKEY_LOCAL_MACHINE" | findstr /i /v !_%arch%code14m! >>"!_temp!\msi.txt"
for %%G in (
"%mvc% 14 %arch% Additional Runtime"
"%mvc% 14 %arch% Minimum Runtime"
"%mvc% 2015 %arch% Additional Runtime"
"%mvc% 2015 %arch% Minimum Runtime"
"%mvc% 2017 %arch% Additional Runtime"
"%mvc% 2017 %arch% Minimum Runtime"
"%mvc% 2019 %arch% Additional Runtime"
"%mvc% 2019 %arch% Minimum Runtime"
) do (
reg query %_natkey% /f %%G /s %_Nul2% | find /i "HKEY_LOCAL_MACHINE" >>"!_temp!\msi.txt"
)

findstr /i "HKEY_LOCAL_MACHINE" "!_temp!\msi.txt" %_Nul3% || goto :process

echo Uninstalling non-compliant Visual C++ MSI packages {%arch%}
echo ^(please wait as this process may take a few moments^)

set invalid=1
for /f "usebackq tokens=7 delims=\" %%G in ("!_temp!\msi.txt") do (
if %_debug% equ 0 (
  start /wait msiexec /X%%G %verbosity% /norestart
  reg delete %_natkey%\%%G /f %_Nul3%
  )
)

:process
set "_x86msi08=2005\x86\vcredist.msi"
set "_x64msi08=2005\x64\vcredist.msi"
set "_x86msi09=2008\x86\vc_red.msi"
set "_x64msi09=2008\x64\vc_red.msi"
set "_x86msi10=2010\x86\vc_red.msi"
set "_x64msi10=2010\x64\vc_red.msi"
set "_x86msi11m=2012\x86\vc_runtimeMinimum_x86.msi"
set "_x86msi11a=2012\x86\vc_runtimeAdditional_x86.msi"
set "_x64msi11m=2012\x64\vc_runtimeMinimum_x64.msi"
set "_x64msi11a=2012\x64\vc_runtimeAdditional_x64.msi"
set "_x86msi12m=2013\x86\vc_runtimeMinimum_x86.msi"
set "_x86msi12a=2013\x86\vc_runtimeAdditional_x86.msi"
set "_x64msi12m=2013\x64\vc_runtimeMinimum_x64.msi"
set "_x64msi12a=2013\x64\vc_runtimeAdditional_x64.msi"
set "_x86msi14m=2022\x86\vc_runtimeMinimum_x86.msi"
set "_x86msi14a=2022\x86\vc_runtimeAdditional_x86.msi"
set "_x64msi14m=2022\x64\vc_runtimeMinimum_x64.msi"
set "_x64msi14a=2022\x64\vc_runtimeAdditional_x64.msi"
set "_x86vstor=vstor\vstor40_x86.msi"
set "_x64vstor=vstor\vstor40_x64.msi"
set "_vbcrun=vbc\vbcrun.msi"

for %%G in (08,09,10) do if !_%arch%install%%G! equ 1 set /a installcount+=1
for %%G in (11,12,14) do if !_%arch%install%%G! equ 1 set /a installcount+=2
if not %arch%==x86 (
for %%G in (08,09,10) do if !_x86install%%G! equ 1 set /a installcount+=1
for %%G in (11,12,14) do if !_x86install%%G! equ 1 set /a installcount+=2
)
if not defined vcpp if !_%arch%installvstor! equ 1 set /a installcount+=1

if %installcount% equ 0 goto :vbc

for %%G in (08,09,10) do if !_%arch%install%%G! equ 1 (
call :install "!_%arch%msi%%G!"
)
for %%G in (11,12,14) do if !_%arch%install%%G! equ 1 (
call :install "!_%arch%msi%%Gm!"
call :install "!_%arch%msi%%Ga!"
)
if not defined vcpp if !_%arch%installvstor! equ 1 (
call :install "!_%arch%vstor!"
)

if %arch%==x86 goto :vbc

for %%G in (08,09,10) do if !_x86install%%G! equ 1 (
call :install "!_x86msi%%G!"
)
for %%G in (11,12,14) do if !_x86install%%G! equ 1 (
call :install "!_x86msi%%Gm!"
call :install "!_x86msi%%Ga!"
)

:vbc
if defined vcpp (
if %installcount% equ 0 if %invalid% equ 0 (call :title&echo All installed Visual C++ Redistributables are compliant.)
goto :close
)
if %arch%==x86 (
set "dest=%SystemRoot%\system32"
set "_qkey=%_natkey%"
) else (
set "dest=%SystemRoot%\syswow64"
set "_qkey=%_wowkey%"
)
if not exist "%dest%\vb40032.dll" goto :vbcinstall
reg query %_qkey%\{C5E3A69D-D391-45A6-A8FB-00B01E2B010D} %_val% %_Nul3% && goto :ucrtbase
for %%G in (
comct232.ocx  msbind.dll    msdbrptr.dll  msstdfmt.dll
comct332.ocx  mscdrun.dll   msflxgrd.ocx  msstkprp.dll
comctl32.ocx  mschrt20.ocx  mshflxgd.ocx  mswcrun.dll
comdlg32.ocx  mscomct2.ocx  mshtmpgr.dll  mswinsck.ocx
dbadapt.dll   mscomctl.ocx  msinet.ocx    picclp32.ocx
dbgrid32.ocx  mscomm32.ocx  msmapi32.ocx  richtx32.ocx
dblist32.ocx  msdatgrd.ocx  msmask32.ocx  sysinfo.ocx
mci32.ocx     msdatlst.ocx  msrdc20.ocx   tabctl32.ocx
msadodc.ocx   msdatrep.ocx  msrdo20.dll
) do (
if exist "%dest%\%%G" set "vbfiles=!vbfiles! %dest%\%%G"
)
if %_debug% equ 0 (%_Nul3% regsvr32 /u /s !vbfiles!) else (echo %vbfiles%)
if %_debug% equ 0 if exist "%dest%\msvbvm50.dll" (
regsvr32 /u /s %dest%\msvbvm50.dll %_Nul3%
reg add HKLM\SYSTEM\CurrentControlSet\Services\EventLog\Application\VBRuntime /v EventMessageFile /t REG_SZ /d %dest%\msvbvm60.dll /f %_Nul3%
reg add HKLM\SYSTEM\CurrentControlSet\Services\EventLog\Application\VBRuntime /v TypesSupported /t REG_DWORD /d 4 /f %_Nul3%
reg import vbc\extra-fix-%arch%.reg %_Nul3%
)
for %%G in (
atl70.dll atl71.dll msvcrt10.dll %dest%\mfc70*.dll %dest%\mfc71*.dll %dest%\msvc*70.dll %dest%\msvc*71.dll
comct232.ocx  msbind.dll    msdbrptr.dll  msstdfmt.dll
comct332.ocx  mscdrun.dll   msflxgrd.ocx  msstkprp.dll
comctl32.ocx  mschrt20.ocx  mshflxgd.ocx  mswcrun.dll
comdlg32.ocx  mscomct2.ocx  mshtmpgr.dll  mswinsck.ocx
dbadapt.dll   mscomctl.ocx  msinet.ocx    picclp32.ocx
dbgrid32.ocx  mscomm32.ocx  msmapi32.ocx  richtx32.ocx
dblist32.ocx  msdatgrd.ocx  msmask32.ocx  sysinfo.ocx
mci32.ocx     msdatlst.ocx  msrdc20.ocx   tabctl32.ocx
msadodc.ocx   msdatrep.ocx  msrdo20.dll   vb40032.dll
) do (
if exist "%dest%\%%~nxG" (
  if %_debug% equ 0 (del /f /q "%dest%\%%~nxG" %_Nul3%)
  )
)
if %_debug% equ 0 if exist "%dest%\msvbvm50.dll" del /f /q %dest%\msvbvm50.dll %_Nul3%
if %_debug% equ 0 if %arch%==x86 (
del /f /q %SystemRoot%\System\vb40016.dll %_Nul3%
del /f /q %SystemRoot%\System\vbrun*.dll %_Nul3%
)

:vbcinstall
call :title
set /a count+=1
set /a installcount+=1
echo Installing %count% of %installcount%: %_vbcrun%
echo.
echo.
if %_debug% equ 0 %_vbcrun% %verbosityshort%

:ucrtbase
if %installcount% equ 0 if %invalid% equ 0 (call :title&echo All installed Visual C++ Redistributables are compliant.)
if defined ucrt goto :close
:ucrtonly
if exist "%SysPath%\ucrtbase.dll" goto :close
if %_debug% equ 1 goto :close
set "dest=%SystemRoot%\servicing\Packages"
if exist "%dest%\Package_for_KB948465*6.0.1.18005.mum" (
start /w PkgMgr.exe /ip /m:"%cd%\ucrt\6002-%arch%.mum" /quiet /norestart %_Nul3%
)
if exist "%dest%\Microsoft-Windows-*6.1.7601.17514.mum" (
dism.exe /Online /Quiet /NoRestart /Add-Package /PackagePath:ucrt\7601-%arch%.mum %_Nul3%
)
if exist "%dest%\Microsoft-Windows-*6.2.9200.16384.mum" (
dism.exe /Online /Quiet /NoRestart /Add-Package /PackagePath:ucrt\9200-%arch%.mum %_Nul3%
)
if exist "%dest%\Package_for_KB2919355*6.3.1.14.mum" (
dism.exe /Online /Quiet /NoRestart /Add-Package /PackagePath:ucrt\9600-%arch%.mum %_Nul3%
)

:close
for %%G in (
"!_temp!\*Redistributable*.*"
"!_temp!\*vcredist*.*"
"!_temp!\*vstor*.*"
"!_temp!\msi*.log"
"!_temp!\del*.tmp"
msi.txt wix.txt
) do (
if exist "!_temp!\%%~nxG" del /f /q "!_temp!\%%~nxG"
)
if exist filever.vbs del /f /q filever.vbs
if defined vbscu reg add "HKCU\%RegKey%" /v Enabled /t REG_DWORD /d 0 /f %_Nul3%
if defined vbslm reg add "HKLM\%RegKey%" /v Enabled /t REG_DWORD /d 0 /f %_Nul3%

if %auto% equ 1 goto :eof
if %installcount% neq 0 (
call :title
echo Installer has completed. 
)
echo.
echo.
echo.
echo Press any key to exit...
pause >nul
goto :eof

:install
call :title
set /a count+=1
echo Installing %count% of %installcount%: %~1
echo.
echo.
if %_debug% equ 0 %~1 %verbosityshort%
goto :eof

:title
if %_debug% equ 1 goto :eof
@cls
echo -----------------------------------------
echo Installer for Visual C++ Runtimes %vci%
echo -----------------------------------------
echo %_os%
echo.
goto :eof

:readme
set _inp=
@cls
echo ----------------------------------------
echo Overview:
echo ----------------------------------------
echo.
echo AIO Repack for latest Microsoft Visual C++ Runtimes
echo without the original setup bloat payload.
echo.
echo Automatically installs the latest available Redistributables for:
echo - Visual C++: 2005, 2008, 2010, 2012, 2013, 2022 ^(2019-2015^)
echo - Visual Studio 2010 Tools for Office Runtime
echo.
echo Additionally, installs these old x86 runtimes:
echo - Visual C++: 2002, 2003
echo - Visual Basic Runtimes
echo.
echo Before installation, the script needs to check and remove
echo existing non-compliant Visual C++ Runtimes.
echo.
echo ----------------------------------------------------------
choice /c NM /N /M "N - Next page, M - Return to Installer"
if errorlevel 2 goto :top
if errorlevel 1 goto :page2
goto :readme

:page2
set _inp=
@cls
echo ----------------------------------------
echo Credits:
echo ----------------------------------------
echo.
echo @ricktendo64 / repacks.net - wincert.net
echo VBCRedist_AIO_x86_x64.exe creator,  modded MSI installers
echo.
echo @burfadel / MDL forums - @thatguy91 / guru3D Forums
echo original installation script
echo.
echo Latest AIO Repack release can be found at:
echo https://tiny.cc/vcredist
echo https://github.com/abbodi1406/vcredist/releases
echo.
echo.
echo Visual Studio is a registered trademark of Microsoft Corporation.
echo.
echo ---------------------------------------------------------------
choice /c PLM /N /M "P - Previous Page, L - Open release link, M - Return to Installer"
if errorlevel 3 goto :top
if errorlevel 2 start https://tiny.cc/vcredist&goto :page2
if errorlevel 1 goto :readme
goto :page2