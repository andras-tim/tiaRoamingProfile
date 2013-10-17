@echo off
taskkill /im explorer.exe /f
start "" wscript //nologo start.vbs start "" "explorer"
exit
