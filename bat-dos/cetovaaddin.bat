echo off


REM ****cetova INSTALL****

if exist c:\windows\cetova_install.txt goto end

echo %computername% >c:\windows\cetova_install.txt

\\forum10\hotfix$\exceladdin-Installer.exe /v/qn

goto end 

:windows

:end



