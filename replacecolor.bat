@echo off

set bookpath=%~dp0sample.xlsx
set target_color=RGB(255,0,0)
set replacement_color=RGB(0,0,0)

cscript //nologo "%~dpn0.vbs" "%bookpath%" "%target_color%" "%replacement_color%"

pause

exit /b 0
