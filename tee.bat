@echo off
pushd "%~dp0"
find /v "" | cscript //nologo "%~dpn0.vbs" %*
