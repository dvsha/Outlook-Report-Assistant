@echo off
cls

SET APPNAME=Outlook Report Assistant 1.0
TITLE %APPNAME%
Echo Installing %APPNAME%
"%~dp0Setup_Files\Setup.vbs"
Exit 0