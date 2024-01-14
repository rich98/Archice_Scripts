@ECHO off
cls
:start
ECHO.
ECHO 1. to dial BTMidband
ECHO 2. to dial vodafone gprs
ECHO 3. to dial bt openworld
set choice=
set /p choice=Type the number to dial.
if not '%choice%'=='' set choice=%choice:~0,1%
if '%choice%'=='1' goto star
if '%choice%'=='2' goto vodafone
if '%choice%'=='3' goto btopenwoprd
ECHO "%choice%" is not valid please try again
ECHO.
goto start
:star
rasdial "BTMidband"
goto check
:vodafone
rasdial "vodafone gprs"
goto check
:btopenworld
rasdial "bt openworld"
goto check

echo not connected!!!

goto start

:check

if exist "\\forum10\hotfix$\laptop user Manual.doc" goto download

echo please check dialup connection and try again 

pause

goto start

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

:end

rasdial /d

pause