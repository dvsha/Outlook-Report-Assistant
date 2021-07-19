@echo off
cls

SET APPNAME=Outlook Report Assistant 1.0
TITLE %APPNAME%
Echo Uninstalling %APPNAME%
rmdir /s /q "%Programdata%\ORA"
rmdir /s /q "%ProgramData%\Microsoft\Windows\Start Menu\Programs\Outlook Report Assistant"

Exit 0