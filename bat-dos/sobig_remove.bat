REM ***** sobig removal *****

echo off

cls
echo **********   Plese read this before you continue    **********
echo **********   search hard drive for Winppr32.exe     **********
echo **********   if this file is found you have Sobig!  **********
echo **********  To cancel close window without running  **********
echo ********** Please backup registry before proceeding!**********

pause

tskill winppr32.exe

del c:\windows\winppr32.exe /f

del c:\windows\winstt32.dat /f

reg delete "hklm\software\Microsoft\Windows\CurrentVersion\Run\TrayX" /f

reg delete "hkcu\software\Microsoft\Windows\CurrentVersion\Run\TrayX" /f

echo ***** run av software ****
pause






