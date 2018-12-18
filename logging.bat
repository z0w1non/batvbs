@echo off

for /f "usebackq tokens=1-7 delims=/:. " %%A in ('%date%%time%') do set datetime=%%A%%B%%C%%D%%E%%F%%G
set log=%~dpn0_%datetime%.log

echo on
help > "%log%"
@echo off

pause
