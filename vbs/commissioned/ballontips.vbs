set Wshshell = WScript.CreateObject("WScript.Shell")


WshShell.RegWrite "HKCU\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced\", "EnableBalloontips"


WshShell.RegWrite "HKCU\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced\EnableBalloontips", "0", "REG_DWORD"