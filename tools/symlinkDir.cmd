@echo off

rem Parse parameters
set argc=0
for %%x in (%*) do set /A argc+=1
if not %argc% == 2 goto error_bad_parameters

rem Check availability of mklink command
mklink 2> "%DEVNULL_ERR%" > "%DEVNULL_OUT%"
if %errorlevel% == 1 goto use_mklink

rem Check ln tool in vendors
if not exist "%trpVENDORS%\ln\ln.exe" goto error_missing_ln_vendor
"%trpVENDORS%\ln\ln.exe" --junction %1 %2 > "%DEVNULL_OUT%"
goto end

:use_mklink
mklink /j %1 %2 > "%DEVNULL_OUT%"
goto end

:error_bad_parameters
echo ERROR: There are not 2 parameters (link, target)
echo.
goto end

:error_missing_ln_vendor
echo ERROR: Missing "ln" tool from "vendors" directory! Please read "placeholder.txt" for details
echo.

:end
