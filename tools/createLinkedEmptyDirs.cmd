@echo off

rem Parse parameters
set argc=0
for %%x in (%*) do set /A argc+=1
if %argc% LSS 2 goto error_bad_parameters

set id=%~1
set link_id=0
for /L %%i in (2,1,%argc%) do call:create_symlink %%%%i
goto end

:create_symlink
set /A link_id+=1
set link_src=%~1
set link_dst=%trpSESSION%\linkedDirs_%id%%link_id%

if not exist "%link_dst%" (
	if exist "%link_src%" rmdir /s /q "%link_src%"
	mkdir "%link_dst%"
	"%trpTOOLS%\symlinkDir.cmd" "%link_src%" "%link_dst%"
)
goto:eof

:error_bad_parameters
echo ERROR: There are'n 2 or more parameters (id, directory1 [, directory2, ...])
echo.

:end
