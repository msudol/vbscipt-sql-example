@echo off
REM RunScript.bat

REM Set the path to your VBScript file
set "vbscriptPath=C:\Path\To\Your\query.vbs"

REM Execute the VBScript
cscript //nologo "%vbscriptPath%"

REM Optional: Pause to keep the command prompt window open (remove if not needed)
pause
