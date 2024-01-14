echo off

if exist "C:\Documents and Settings\%username%\Application Data\Microsoft\Templates\Bioscience.pot" goto end

pause

:Bio_template

xcopy "\\forum10\hotfix$\bioscience.pot" "C:\Documents and Settings\%username%\Application Data\Microsoft\Templates"

:end

pause
