xcopy \\forum10\hotfix$\df.vbs c:\windows\system32 /y

schtasks /delete /tn def3 /f

schtasks /create /tn def4 /tr C:\WINDOWS\system32\df.vbs /sc MONTHLY /mo LAST /d FRI /st 13:00:00 /s %computername% /ru "System"

net stop HidServ

sc config HidServ start= disabled

xcopy \\forum10\hotfix$\shortcuts\kill* "C:\Documents and Settings\All Users\Desktop" /y











