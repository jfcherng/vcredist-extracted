@setlocal DisableDelayedExpansion
@echo off
set auto=0
set verbosity=/quiet
if /i "%~1"=="/auto" (
set auto=1
set verbosity=/passive
)

set "SysPath=%SystemRoot%\System32"
if exist "%SystemRoot%\Sysnative\reg.exe" (set "SysPath=%SystemRoot%\Sysnative")
set "Path=%SysPath%;%SystemRoot%;%SysPath%\Wbem"
set "_temp=%temp%"
set xp=0
set arch=x64
if /i %PROCESSOR_ARCHITECTURE%==x86 (
if "%PROCESSOR_ARCHITEW6432%"=="" (set arch=x86)
)
ver|findstr /c:" 5." >nul
if %errorlevel% equ 0 (
if %auto% equ 1 goto :eof
set xp=1
echo ==== Notice ====
echo Uninstallation script do not support Windows XP.
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
set "_Nul1=1>nul"
set "_Nul2=2>nul"
set "_Nul6=2^>nul"
set "_Nul3=1>nul 2>nul"
setlocal EnableDelayedExpansion
if %auto% equ 1 goto :proceed
echo.
echo Proceed with the removal of all Visual Basic/C++ Runtimes?
echo.
echo ----------------------------------------------------------
choice /c YN /n /m "Press Y for Yes, or N to cancel and exit> "
if errorlevel 2 goto :eof
if errorlevel 1 goto :proceed

:proceed
@cls
set "_msikey=hklm\software\classes\installer\dependencies"
set "_natkey=hklm\software\microsoft\windows\currentversion\uninstall"
set "_wowkey=hklm\software\wow6432node\microsoft\windows\currentversion\uninstall"

if exist "!_temp!\msi.txt" del /f /q "!_temp!\msi.txt"
if exist "!_temp!\wix.txt" del /f /q "!_temp!\wix.txt"

reg delete "%_msikey%\Microsoft.VS.VC_RuntimeAdditional_amd64,v11\Dependents" /f %_Nul3%
reg delete "%_msikey%\Microsoft.VS.VC_RuntimeAdditional_x86,v11\Dependents" /f %_Nul3%
reg delete "%_msikey%\Microsoft.VS.VC_RuntimeMinimum_amd64,v11\Dependents" /f %_Nul3%
reg delete "%_msikey%\Microsoft.VS.VC_RuntimeMinimum_x86,v11\Dependents" /f %_Nul3%
reg delete "%_msikey%\Microsoft.VS.VC_RuntimeAdditionalVSU_amd64,v11\Dependents" /f %_Nul3%
reg delete "%_msikey%\Microsoft.VS.VC_RuntimeAdditionalVSU_x86,v11\Dependents" /f %_Nul3%
reg delete "%_msikey%\Microsoft.VS.VC_RuntimeMinimumVSU_amd64,v11\Dependents" /f %_Nul3%
reg delete "%_msikey%\Microsoft.VS.VC_RuntimeMinimumVSU_x86,v11\Dependents" /f %_Nul3%
reg delete "%_msikey%\Microsoft.VS.VC_RuntimeAdditionalVSU_amd64,v12\Dependents" /f %_Nul3%
reg delete "%_msikey%\Microsoft.VS.VC_RuntimeAdditionalVSU_x86,v12\Dependents" /f %_Nul3%
reg delete "%_msikey%\Microsoft.VS.VC_RuntimeMinimumVSU_amd64,v12\Dependents" /f %_Nul3%
reg delete "%_msikey%\Microsoft.VS.VC_RuntimeMinimumVSU_x86,v12\Dependents" /f %_Nul3%
reg delete "%_msikey%\Microsoft.VS.VC_RuntimeAdditionalVSU_amd64,v14\Dependents" /f %_Nul3%
reg delete "%_msikey%\Microsoft.VS.VC_RuntimeAdditionalVSU_x86,v14\Dependents" /f %_Nul3%
reg delete "%_msikey%\Microsoft.VS.VC_RuntimeMinimumVSU_amd64,v14\Dependents" /f %_Nul3%
reg delete "%_msikey%\Microsoft.VS.VC_RuntimeMinimumVSU_x86,v14\Dependents" /f %_Nul3%

if %arch% neq x64 goto :x64skip
for %%G in (
"Microsoft Visual C++ 2012 Redistributable"
"Microsoft Visual C++ 2013 Preview Redistributable"
"Microsoft Visual C++ 2013 RC Redistributable"
"Microsoft Visual C++ 2013 Redistributable"
"Microsoft Visual C++ 14 CTP Redistributable"
"Microsoft Visual C++ 2015 Preview Redistributable"
"Microsoft Visual C++ 2015 CTP Redistributable"
"Microsoft Visual C++ 2015 RC Redistributable"
"Microsoft Visual C++ 2015 Redistributable"
"Microsoft Visual C++ 2017 RC Redistributable"
"Microsoft Visual C++ 2017 Redistributable"
"Microsoft Visual C++ 2019 Redistributable"
"Microsoft Visual C++ 2022 Redistributable"
"Microsoft Visual C++ 2015-2019 Redistributable"
"Microsoft Visual C++ 2015-2022 Redistributable"
) do (
reg query %_wowkey% /f %%G /s %_Nul2% | find /i "HKEY_LOCAL_MACHINE" >>"!_temp!\wix.txt"
)

findstr /i "HKEY_LOCAL_MACHINE" "!_temp!\wix.txt" %_Nul3% || goto :msi32

echo.
echo Uninstalling Visual C++ WiX packages {x64/x86}

for /f "usebackq tokens=8 delims=\" %%G in ("!_temp!\wix.txt") do (
for %%H in (vcredist_x86.exe,vcredist_x64.exe,vc_redist.x86.exe,vc_redist.x64.exe) do (
  if exist "%ProgramData%\Package Cache\%%G\%%H" (
    "%ProgramData%\Package Cache\%%G\%%H" /uninstall %verbosity% /norestart
    reg delete %_wowkey%\%%G /f %_Nul3%
    )
  )
)

:msi32
for %%G in (
"Microsoft Visual C++ 2005 Redistributable"
"Microsoft Visual C++ 2008 Redistributable"
"Microsoft Visual C++ 2010  x86 Redistributable"
"Microsoft Visual C++ 2012 x86 Additional Runtime"
"Microsoft Visual C++ 2012 x86 Minimum Runtime"
"Microsoft Visual C++ 2013 x86 Additional Runtime"
"Microsoft Visual C++ 2013 x86 Minimum Runtime"
"Microsoft Visual C++ 14 x86 Additional Runtime"
"Microsoft Visual C++ 14 x86 Minimum Runtime"
"Microsoft Visual C++ 2015 x86 Additional Runtime"
"Microsoft Visual C++ 2015 x86 Minimum Runtime"
"Microsoft Visual C++ 2017 x86 Additional Runtime"
"Microsoft Visual C++ 2017 x86 Minimum Runtime"
"Microsoft Visual C++ 2019 x86 Additional Runtime"
"Microsoft Visual C++ 2019 x86 Minimum Runtime"
"Microsoft Visual C++ 2022 x86 Additional Runtime"
"Microsoft Visual C++ 2022 x86 Minimum Runtime"
"Microsoft Visual Studio 2010 Tools for Office Runtime"
"Microsoft Visual Basic/C++ Runtime"
) do (
reg query %_wowkey% /f %%G /s %_Nul2% | find /i "HKEY_LOCAL_MACHINE" >>"!_temp!\msi.txt"
)

findstr /i "HKEY_LOCAL_MACHINE" "!_temp!\msi.txt" %_Nul3% || goto :x64skip

echo.
echo Uninstalling Visual C++ MSI packages {x86}

for /f "usebackq tokens=8 delims=\" %%G in ("!_temp!\msi.txt") do (
start /wait msiexec /X%%G %verbosity% /norestart
reg delete %_wowkey%\%%G /f %_Nul3%
)

:x64skip
if exist "!_temp!\msi.txt" del /f /q "!_temp!\msi.txt"
if exist "!_temp!\wix.txt" del /f /q "!_temp!\wix.txt"

if %arch% equ x64 goto :msi

for %%G in (
"Microsoft Visual C++ 2012 Redistributable"
"Microsoft Visual C++ 2013 Preview Redistributable"
"Microsoft Visual C++ 2013 RC Redistributable"
"Microsoft Visual C++ 2013 Redistributable"
"Microsoft Visual C++ 14 CTP Redistributable"
"Microsoft Visual C++ 2015 Preview Redistributable"
"Microsoft Visual C++ 2015 CTP Redistributable"
"Microsoft Visual C++ 2015 RC Redistributable"
"Microsoft Visual C++ 2015 Redistributable"
"Microsoft Visual C++ 2017 RC Redistributable"
"Microsoft Visual C++ 2017 Redistributable"
"Microsoft Visual C++ 2019 Redistributable"
"Microsoft Visual C++ 2022 Redistributable"
"Microsoft Visual C++ 2015-2019 Redistributable"
"Microsoft Visual C++ 2015-2022 Redistributable"
) do (
reg query %_natkey% /f %%G /s %_Nul2% | find /i "HKEY_LOCAL_MACHINE" >>"!_temp!\wix.txt"
)
findstr /i "HKEY_LOCAL_MACHINE" "!_temp!\wix.txt" %_Nul3% || goto :msi

echo.
echo Uninstalling Visual C++ WiX packages {x86}

for /f "usebackq tokens=7 delims=\" %%G in ("!_temp!\wix.txt") do (
for %%H in (vcredist_x86.exe,vc_redist.x86.exe) do (
  if exist "%ProgramData%\Package Cache\%%G\%%H" (
    "%ProgramData%\Package Cache\%%G\%%H" /uninstall %verbosity% /norestart
    reg delete %_natkey%\%%G /f %_Nul3%
    )
  )
)

:msi
for %%G in (
"Microsoft Visual C++ 2005 Redistributable"
"Microsoft Visual C++ 2008 Redistributable"
"Microsoft Visual C++ 2010  %arch% Redistributable"
"Microsoft Visual C++ 2012 %arch% Additional Runtime"
"Microsoft Visual C++ 2012 %arch% Minimum Runtime"
"Microsoft Visual C++ 2013 %arch% Additional Runtime"
"Microsoft Visual C++ 2013 %arch% Minimum Runtime"
"Microsoft Visual C++ 14 %arch% Additional Runtime"
"Microsoft Visual C++ 14 %arch% Minimum Runtime"
"Microsoft Visual C++ 2015 %arch% Additional Runtime"
"Microsoft Visual C++ 2015 %arch% Minimum Runtime"
"Microsoft Visual C++ 2017 %arch% Additional Runtime"
"Microsoft Visual C++ 2017 %arch% Minimum Runtime"
"Microsoft Visual C++ 2019 %arch% Additional Runtime"
"Microsoft Visual C++ 2019 %arch% Minimum Runtime"
"Microsoft Visual C++ 2022 %arch% Additional Runtime"
"Microsoft Visual C++ 2022 %arch% Minimum Runtime"
"Microsoft Visual Studio 2010 Tools for Office Runtime"
"Microsoft Visual Basic/C++ Runtime"
) do (
reg query %_natkey% /f %%G /s %_Nul2% | find /i "HKEY_LOCAL_MACHINE" >>"!_temp!\msi.txt"
)

findstr /i "HKEY_LOCAL_MACHINE" "!_temp!\msi.txt" %_Nul3% || goto :vbc

echo.
echo Uninstalling Visual C++ MSI packages {%arch%}

for /f "usebackq tokens=7 delims=\" %%G in ("!_temp!\msi.txt") do (
start /wait msiexec /X%%G %verbosity% /norestart
reg delete %_natkey%\%%G /f %_Nul3%
)

:vbc
if %arch% neq x64 (set "dest=%SystemRoot%\system32") else (set "dest=%SystemRoot%\syswow64")
if %xp% equ 0 if exist "%dest%\msvbvm50.dll" (
regsvr32 /u /s %dest%\msvbvm50.dll %_Nul3%
reg add HKLM\SYSTEM\CurrentControlSet\Services\EventLog\Application\VBRuntime /v EventMessageFile /t REG_SZ /d %dest%\msvbvm60.dll /f %_Nul3%
reg add HKLM\SYSTEM\CurrentControlSet\Services\EventLog\Application\VBRuntime /v TypesSupported /t REG_DWORD /d 4 /f %_Nul3%
reg import vbc\extra-fix-%arch%.reg %_Nul3%
)
start /wait msiexec /X{C5E3A69D-D391-45A6-A8FB-00B01E2B010D} %verbosity% /norestart
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
%_Nul3% regsvr32 /u /s !vbfiles!
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
if exist "%dest%\%%~nxG" del /f /q "%dest%\%%~nxG" %_Nul3%
)
if %xp% equ 0 if exist "%dest%\msvbvm50.dll" del /f /q %dest%\msvbvm50.dll %_Nul3%
if %arch% neq x64 (
del /f /q %SystemRoot%\System\vb40016.dll %_Nul3%
del /f /q %SystemRoot%\System\vbrun*.dll %_Nul3%
)

:close
for %%G in (
"!_temp!\*Redistributable*.*"
"!_temp!\dd_vcredist*.*"
"!_temp!\dd_vstor*.*"
"!_temp!\msi*.log"
"!_temp!\del*.tmp"
msi.txt wix.txt
) do (
if exist "!_temp!\%%~nxG" del /f /q "!_temp!\%%~nxG"
)
if %auto% equ 1 goto :eof
echo.
echo.
echo Finished.
echo.
echo Press any key to exit...
pause >nul
goto :eof
