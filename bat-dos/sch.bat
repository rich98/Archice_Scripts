schtasks /create /tn defrag /tr C:\WINDOWS\system32\defrag.exe /sc onidle /i 10 /s %computername% /ru "system"


