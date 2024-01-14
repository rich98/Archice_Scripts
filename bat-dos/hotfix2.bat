echo off

REM**** hotfix install second bat ****

if exist c:\windows\%computername%2.txt goto end

if exist c:\%computername%2.txt goto end


REM \\matrix2000\hotfix$\Q328310.exe -q -z


\\matrix2000\hotfix$\q813489.exe -q -z


\\matrix2000\hotfix$\q329115.exe -q -z


\\matrix2000\hotfix$\q329170.exe -q -z


\\matrix2000\hotfix$\q810577.exe -q -z


\\matrix2000\hotfix$\q331953.exe -q -z

echo %computername%2 >c:\%computername%2.txt

xcopy c:\windows\%computername%2.txt \\matrix2000\hotfix$

xcopy c:\%computername%2.txt \\matrix2000\hotfix$ 


:end







