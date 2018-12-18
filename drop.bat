@echo off

rem Drop the one or multiple files to this bat file.

set vbs=%~dpn0.vbs

:for_each
cscript //nologo "%vbs%" "%~1"
shift
if not "%~1"=="" goto for_each

pause

exit /b 0
