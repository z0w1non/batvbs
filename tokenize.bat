@echo off

pushd "%~dp0"
set out=%~dpn0.txt
call :split > "%out%"
exit /b 0

:split
    for /f "usebackq delims=" %%I in (`dir /b /s /a-d ^| find /v "%out%"`) do (
        cscript //nologo "%~dpn0.vbs" < "%%I"
    )
exit /b 0
