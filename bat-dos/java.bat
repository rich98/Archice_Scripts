echo off

REM **** Java & update ****

if exist c:\windows\%computername%3.txt goto end

if exist c:\%computername%3.txt goto end 

\\matrix2000\hotfix$\js56nen.exe -q
 

\\matrix2000\hotfix$\q811493.exe -q -z


REM \\matrix2000\hofix$\wmupdate.exe -q 


echo java update >c:\windows\%computername%3.txt

xcopy c:\windows\%computername%3.txt \\matrix2000\hotfix$

:end

exit






