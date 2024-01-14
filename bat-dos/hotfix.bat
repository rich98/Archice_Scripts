echo off


REM ****HOTFIX INSTALL****

if exist c:\windows\cetova_install.tx goto end

echo %computername% >c:\windows\cetova_install.txt

\\forum10\hotfix$\exceladdin-Installer.exe /s

goto end 

:windows

:end



