echo off


REM ****laptop user doc****

del "C:\Documents and Settings\All Users\Desktop\Laptop User Manual v1.doc"


xcopy "\\forum10\hotfix$\Laptop User Manual.doc" "C:\Documents and Settings\All Users\Desktop" /y


pause
:end


