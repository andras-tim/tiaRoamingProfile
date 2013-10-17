@echo off

rem Parse parameters
set argc=0
for %%x in (%*) do set /A argc+=1
if %argc% == 0 goto error_missing

set command_params=
for /L %%i in (2,1,%argc%) do call set command_params=%%command_params%% %%%%i

rem Choose Program Files folder
set command=NULL
if exist "%ProgramFiles%\%~1" set command=%ProgramFiles%\%~1
if exist "%ProgramFiles(x86)%\%~1" set command=%ProgramFiles(x86)%\%~1
if "%command%" == "NULL" goto error_not_exists

rem Run command with parameters
start "" "%command%" %command_params%
goto end

:error_missing
echo ERROR: Missing path of program from parameter
echo.
goto end

:error_not_exists
echo ERROR: Not exists the requested path under Program Files folder(s) - if exists Program Files (x86)
echo.

:end