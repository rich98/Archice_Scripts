echo off


REM ****laptop user doc****

rasdial "star" 

ping 10.9.150.22 

if exist "\\forum10\hotfix$\laptop user Manual.doc" goto download

echo please check dialup connection and try again 

pause

goto end

:download

cls

echo You are about to download the "laptop user manual"

pause

echo downloading............Please wait..........

echo off

xcopy "\\forum10\hotfix$\Laptop User Manual.doc" "C:\Documents and Settings\All Users\Desktop" /y

del "C:\Documents and Settings\All Users\Desktop\Laptop User Manual v1.doc"

color 47

cls

echo finished Thank you

rasdial "star" /d
pause

:end

rasdial "star" /d