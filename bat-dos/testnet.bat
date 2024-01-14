net stop Mcshield
pause

REM del "C:\Program Files\Network Associates\VirusScan\*.mmf" 
pause

regedit.exe /s \\matrix2000\avregfiles$\NETAV.REG
pause

net start McShield
pause