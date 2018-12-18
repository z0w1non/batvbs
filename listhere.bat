@echo off

pushd "%~dp0"

echo *** Files and directories
for /f "usebackq delims=" %%I in (`dir /b`) do call :for_each "%%I"
echo.

echo *** Only files
for /f "usebackq delims=" %%I in (`dir /b /a-d`) do call :for_each "%%I"
echo.

echo *** Only directories
for /f "usebackq delims=" %%I in (`dir /b /ad`) do call :for_each "%%I"
echo.

pause
exit /b 0

:for_each
    echo %~1
exit /b 0
