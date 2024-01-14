REM ***** IT@Forum install tools *****

REM if exist c:\windows\toolinstalled.txt goto end

mkdir "C:\Documents and Settings\All Users\Start Menu\Programs\it@forum"

xcopy \\matrix2000\hotfix$\tools\*.* c:\windows /y

xcopy \\matrix2000\hotfix$\shortcuts\*.* "C:\Documents and Settings\All Users\Start Menu\Programs\it@forum" /y

echo IT@forum tools installed >c:\windows\toolsinstalled.txt

:end



