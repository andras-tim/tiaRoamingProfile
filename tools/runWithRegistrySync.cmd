@echo off

rem Parse parameters
set argc=0
for %%x in (%*) do set /A argc+=1
if %argc% == 0 goto error_missing_prg

set command=%~1
set command_params=
for /L %%i in (2,1,%argc%) do set command_params=%%command_params%% %%%%i

set basename=
for /F %%i in ("%command%") do set basename=%%~ni

rem Check registry key
if %REG_KEY% == "" goto error_not_set_reg_key

rem Initialize environment
set NULL_OUT=%TEMP%\tmp_stdout
set NULL_ERR=%TEMP%\tmp_stderr
if not exist "%trpDATA%\REGCONF" mkdir "%trpDATA%\REGCONF"
set conf_path=%trpDATA%\REGCONF\%basename%.reg

rem Cleanup registry
reg delete %REG_KEY% /f > "%NULL_OUT%" 2> "%NULL_ERR%"

rem Restore registry
if exist "%conf_path%" reg import "%conf_path%" > "%NULL_OUT%" 2> "%NULL_ERR%"

rem Run command with parameters
call "%command%" %command_params%

rem Backup registry
reg export %REG_KEY% "%conf_path%" /y > "%NULL_OUT%" 2> "%NULL_ERR%"

rem Cleanup registry
reg delete %REG_KEY% /f > "%NULL_OUT%" 2> "%NULL_ERR%"
goto end

:error_missing_prg
echo ERROR: Missing path of program from parameter
echo.
goto end

:error_not_set_reg_key
echo ERROR: Not set REG_KEY environment variable, what want to represent key of application settings
echo.

:end