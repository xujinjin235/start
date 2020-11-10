@echo off

:: BatchGotAdmin
:-------------------------------------
REM  --> Check for permissions
>nul 2>&1 "%SYSTEMROOT%\system32\cacls.exe" "%SYSTEMROOT%\system32\config\system"

REM --> If error flag set, we do not have admin.
if '%errorlevel%' NEQ '0' (
    echo Requesting administrative privileges...
    goto UACPrompt
) else ( goto gotAdmin )

:UACPrompt
    echo Set UAC = CreateObject^("Shell.Application"^) > "%temp%\getadmin.vbs"
    set params = %*:"=""
    echo UAC.ShellExecute "cmd.exe", "/c %~s0 %params%", "", "runas", 1 >> "%temp%\getadmin.vbs"

    "%temp%\getadmin.vbs"
    del "%temp%\getadmin.vbs"
    exit /B

:gotAdmin
    pushd "%CD%"
    CD /D "%~dp0"
:--------------------------------------

color 2F
echo.
echo.
echo.1.Office 2010 激活
echo.
echo.2.Office 2013 激活
echo.
echo.3.Office 2016 激活
echo.
echo.4.Windows 激活
echo.
echo.
set KMS_Server=210.34.0.137
goto 3
:office
setlocal EnableDelayedExpansion
reg query %strRegKey% >nul 2>nul
if %errorlevel%==0 (set strCurrentKey=%strRegKey%) else (set strCurrentKey=%strRegKey6432%)
for /f "delims=" %%i in ('reg query %strCurrentKey%') do (
set strInstPath=%%i
set strInstPath=!strInstPath:*REG_SZ=!
)
:LTrim
if "%strInstPath:~0,1%"==" " set "strInstPath=%strInstPath:~1%" && goto LTrim
:RTrim
if "%strInstPath:~-1%"==" " set "strInstPath=%strInstPath:~0,-1%" && goto RTrim
if "%strInstPath:~-1%" neq "\" set strInstPath=%strInstPath%\
echo office安装目录为%strInstPath% 
cd /d %strInstPath%
cscript ospp.vbs /sethst:%KMS_Server%
cscript ospp.vbs /act
pause
exit

:1
set "strRegKey=HKEY_LOCAL_MACHINE\Software\Microsoft\Office\14.0\Common\InstallRoot /v Path"
set "strRegKey6432=HKEY_LOCAL_MACHINE\Software\Wow6432Node\Microsoft\Office\14.0\Common\InstallRoot /v Path"
goto office

:2
set "strRegKey=HKEY_LOCAL_MACHINE\Software\Microsoft\Office\15.0\Common\InstallRoot /v Path"
set "strRegKey6432=HKEY_LOCAL_MACHINE\Software\Wow6432Node\Microsoft\Office\15.0\Common\InstallRoot /v Path"
goto office

:3
set "strRegKey=HKEY_LOCAL_MACHINE\Software\Microsoft\Office\16.0\Common\InstallRoot /v Path"
set "strRegKey6432=HKEY_LOCAL_MACHINE\Software\Wow6432Node\Microsoft\Office\16.0\Common\InstallRoot /v Path"
goto office


:4
cscript "%SystemRoot%\system32\slmgr.vbs" /skms %KMS_Server%
cscript "%SystemRoot%\system32\slmgr.vbs" -ato
pause
exit