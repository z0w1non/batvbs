@echo off

set word=FileSystemObject
set target=%~dp0

cscript //nologo "%~dpn0.vbs" "%word%" "%target%"
pause

exit /b 0
