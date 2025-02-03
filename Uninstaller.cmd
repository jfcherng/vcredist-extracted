@setlocal DisableDelayedExpansion
@echo off
set verbosity=/quiet
set auto=0
set bt32=0
if /i "%~1"=="/auto" (
set auto=1
set verbosity=/passive
)
if /i "%~2"=="/x86" set bt32=1

set "SysPath=%SystemRoot%\System32"
set "Path=%SystemRoot%\System32;%SystemRoot%;%SystemRoot%\System32\Wbem;%SystemRoot%\System32\WindowsPowerShell\v1.0\"
if exist "%SystemRoot%\Sysnative\reg.exe" (
set "SysPath=%SystemRoot%\Sysnative"
set "Path=%SystemRoot%\Sysnative;%SystemRoot%;%SystemRoot%\Sysnative\Wbem;%SystemRoot%\Sysnative\WindowsPowerShell\v1.0\;%Path%"
)
set "_temp=%temp%"

set "arch=x64"
if /i "%PROCESSOR_ARCHITECTURE%"=="x86" if "%PROCESSOR_ARCHITEW6432%"=="" set "arch=x86"
if /i "%PROCESSOR_ARCHITECTURE%"=="arm64" set "arch=arm64"
if /i "%PROCESSOR_ARCHITEW6432%"=="arm64" set "arch=arm64"
if /i "%PROCESSOR_ARCHITEW6432%"=="amd64" set "arch=x64"
set wixpkg=vcredist_x86.exe,vc_redist.x86.exe,vcredist_x64.exe,vc_redist.x64.exe
if %arch%==arm64 set wixpkg=vcredist_x86.exe,vc_redist.x86.exe

set _xp=0
REG QUERY "HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion" /v CurrentBuildNumber |FINDSTR 2600 >NUL
if %errorlevel% equ 0 set _xp=1
if %_xp% equ 1 goto :unsupported

reg.exe query "HKU\S-1-5-19" 1>nul 2>nul || goto :unadmin

set "_f_=comct232.ocx msbind.dll msdbrptr.dll msstdfmt.dll comct332.ocx mscdrun.dll msflxgrd.ocx msstkprp.dll comctl32.ocx mschrt20.ocx mshflxgd.ocx mswcrun.dll comdlg32.ocx mscomct2.ocx mshtmpgr.dll mswinsck.ocx dbadapt.dll mscomctl.ocx msinet.ocx picclp32.ocx dbgrid32.ocx mscomm32.ocx msmapi32.ocx richtx32.ocx dblist32.ocx msdatgrd.ocx msmask32.ocx sysinfo.ocx mci32.ocx msdatlst.ocx msrdc20.ocx tabctl32.ocx msadodc.ocx msdatrep.ocx msrdo20.dll"

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
if %bt32% equ 1 set wixpkg=vcredist_x86.exe,vc_redist.x86.exe
set "mvc=Microsoft Visual C++"
set "mvt=Microsoft Visual Studio 2010 Tools for Office Runtime"
set "_g_="{.*-.*-.*-.*-.*}""
set "_h_="HKEY_LOCAL_MACHINE""
set "_r_=Redistributable"
set "_l_=Additional Runtime"
set "_m_=Minimum Runtime"
set "_natkey=HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall"
set "_wowkey=HKLM\SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall"
set "_msikey=HKLM\SOFTWARE\Classes\Installer\Dependencies"

if %arch%==x86 (
call :unDEP %arch%
) else (
call :unDEP x86
if %bt32% equ 0 call :unDEP %arch%
)

:WiXWow
if exist "!_temp!\msi.txt" del /f /q "!_temp!\msi.txt"
if exist "!_temp!\wix.txt" del /f /q "!_temp!\wix.txt"

if %arch%==x86 goto :WiXNat

call :chWiX wow

findstr /i %_h_% "!_temp!\wix.txt" %_Nul3% || goto :MsiWow

call :unWiX wow

:MsiWow
call :chMSI wow

findstr /i %_h_% "!_temp!\msi.txt" %_Nul3% || goto :WiXNat

call :unMSI wow

:WiXNat
if exist "!_temp!\msi.txt" del /f /q "!_temp!\msi.txt"
if exist "!_temp!\wix.txt" del /f /q "!_temp!\wix.txt"

if not %arch%==x86 goto :MsiNat

call :chWiX nat

findstr /i %_h_% "!_temp!\wix.txt" %_Nul3% || goto :MsiNat

call :unWiX nat

:MsiNat
call :chMSI nat

findstr /i %_h_% "!_temp!\msi.txt" %_Nul3% || goto :chVBC

call :unMSI nat

:chVBC
if %arch%==x86 (
set "dest=%SystemRoot%\system32"
set "_cpf=%CommonProgramFiles%\DESIGNER"
) else (
set "dest=%SystemRoot%\syswow64"
set "_cpf=%CommonProgramFiles(x86)%\DESIGNER"
)
if %_xp% equ 0 if exist "%dest%\msvbvm50.dll" (
regsvr32 /u /s %dest%\msvbvm50.dll %_Nul3%
reg add HKLM\SYSTEM\CurrentControlSet\Services\EventLog\Application\VBRuntime /v EventMessageFile /t REG_SZ /d %dest%\msvbvm60.dll /f %_Nul3%
reg add HKLM\SYSTEM\CurrentControlSet\Services\EventLog\Application\VBRuntime /v TypesSupported /t REG_DWORD /d 4 /f %_Nul3%
reg import vbc\extra-fix-%arch%.reg %_Nul3%
)
start /wait MsiExec.exe /X{C5E3A69D-D391-45A6-A8FB-00B01E2B010D} %verbosity% /norestart
start /wait MsiExec.exe /X{C5E3A69D-D392-45A6-A8FB-00B01E2B010D} %verbosity% /norestart
start /wait MsiExec.exe /X{C5E3A69D-D393-45A6-A8FB-00B01E2B010D} %verbosity% /norestart
for %%G in (
%_f_%
) do if exist "%dest%\%%G" (
%_Nul3% regsvr32 /u /s %dest%\%%G
)
for %%G in (
mscdrun.dll mshtmpgr.dll mswcrun.dll
) do if exist "%_cpf%\%%G" (
%_Nul3% regsvr32 /u /s "%_cpf%\%%G"
)
for %%G in (
atl70.dll atl71.dll msvcrt10.dll vb40032.dll %dest%\mfc70*.dll %dest%\mfc71*.dll %dest%\msvc*70.dll %dest%\msvc*71.dll
%_f_%
) do if exist "%dest%\%%~nxG" (
del /f /q "%dest%\%%~nxG" %_Nul3%
)
for %%G in (
mscdrun.dll mshtmpgr.dll mswcrun.dll
) do if exist "%_cpf%\%%G" (
del /f /q "%_cpf%\%%~nxG" %_Nul3%
)
if %_xp% equ 0 if exist "%dest%\msvbvm50.dll" del /f /q %dest%\msvbvm50.dll %_Nul3%
if %arch%==x86 (
del /f /q %SystemRoot%\System\vb40016.dll %_Nul3%
del /f /q %SystemRoot%\System\vbrun*.dll %_Nul3%
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
if %auto% equ 1 goto :eof
echo.
echo.
echo Finished.
goto :E_Exit

:unadmin
echo ==== ERROR ====
echo This script require administrator privileges.
echo To do so, right click on this script and select 'Run as administrator'
if %auto% equ 1 goto :eof
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

:unDEP
set _a_=%1
for %%G in (v11,v12,v14) do (
reg delete "%_msikey%\Microsoft.VS.VC_RuntimeAdditional_%_a_%,%%G\Dependents" /f %_Nul3%
reg delete "%_msikey%\Microsoft.VS.VC_RuntimeMinimum_%_a_%,%%G\Dependents" /f %_Nul3%
reg delete "%_msikey%\Microsoft.VS.VC_RuntimeAdditionalVSU_%_a_%,%%G\Dependents" /f %_Nul3%
reg delete "%_msikey%\Microsoft.VS.VC_RuntimeMinimumVSU_%_a_%,%%G\Dependents" /f %_Nul3%
)
goto :eof

:chWiX
if "%1"=="nat" (
set _k_=%_natkey%
) else (
set _k_=%_wowkey%
)

for %%G in (
"%mvc% 2012 %_r_%"
"%mvc% 2013 Preview %_r_%"
"%mvc% 2013 RC %_r_%"
"%mvc% 2013 %_r_%"
"%mvc% 14 CTP %_r_%"
"%mvc% 2015 Preview %_r_%"
"%mvc% 2015 CTP %_r_%"
"%mvc% 2015 RC %_r_%"
"%mvc% 2015 %_r_%"
"%mvc% 2017 RC %_r_%"
"%mvc% 2017 %_r_%"
"%mvc% 2019 %_r_%"
"%mvc% 2022 %_r_%"
"%mvc% 2015-2019 %_r_%"
"%mvc% 2015-2022 %_r_%"
) do (
reg query %_k_% /f %%G /s %_Nul2% | find /i %_h_% | findstr /r %_g_% >>"!_temp!\wix.txt"
)

goto :eof

:chMSI
if "%1"=="nat" (
set _a_=%arch%
set _k_=%_natkey%
) else (
set _a_=x86
set _k_=%_wowkey%
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
reg query %_k_% /f %%G /s %_Nul2% | find /i %_h_% | findstr /r %_g_% >>"!_temp!\msi.txt"
)

goto :eof

:unWiX
if "%1"=="nat" (
set _k_=%_natkey%
set _t_=x86
set _c_=7
set _p_=vcredist_x86.exe,vc_redist.x86.exe
) else (
set _k_=%_wowkey%
set _t_=x64/x86
set _c_=8
set _p_=%wixpkg%
)

echo.
echo Uninstalling Visual C++ WiX packages {%_t_%}

for /f "usebackq tokens=%_c_% delims=\" %%G in ("!_temp!\wix.txt") do (
for %%H in (%_p_%) do (
  if exist "%ProgramData%\Package Cache\%%G\%%H" (
      start /wait "" "%ProgramData%\Package Cache\%%G\%%H" /uninstall %verbosity% /norestart
      reg delete %_k_%\%%G /f %_Nul3%
    )
  )
)

goto :eof

:unMSI
if "%1"=="nat" (
set _k_=%_natkey%
set _t_=%arch%
set _c_=7
) else (
set _k_=%_wowkey%
set _t_=x86
set _c_=8
)

echo.
echo Uninstalling Visual C++ MSI packages {%_t_%}

for /f "usebackq tokens=%_c_% delims=\" %%G in ("!_temp!\msi.txt") do (
  start /wait MsiExec.exe /X%%G %verbosity% /norestart
  reg delete %_k_%\%%G /f %_Nul3%
)

goto :eof
