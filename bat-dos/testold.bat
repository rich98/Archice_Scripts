echo off
REM ****test alert****
echo off
color 4f

echo Installing Hotfix's please wait do not close this window Thank You...........

pause
REM ****HOTFIX INSTALL****

echo Now running Q328145

echo off

\\matrix2000\hotfix$\Q328145.exe -q -z

echo Now running Q323255

echo off

\\matrix2000\hotfix$\Q323255.exe -q -z

echo Now running Q329834

echo off

\\matrix2000\hotfix$\Q329834.exe -q -z

echo Now Running Q329390

echo off

\\matrix2000\hotfix$\Q329390.exe -q -z


echo Now Running Q810833

echo off

\\matrix2000\hotfix$\Q810833.exe -q -z

echo Now Running Q329048

echo off

\\matrix2000\hotfix$\Q329048.exe -q -z

echo Now Running Q328310 Please wait..........

echo off
 
\\matrix2000\hotfix$\Q328310.exe -q -z

echo %computername%> %computername%.txt

xcopy c:\%computername%.txt \\matrix2000\hotfix$

del c:\%computername%.txt
cls
color 1F

echo Hotfix's installed Thank You For your Time

pause