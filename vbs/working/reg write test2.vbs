set Wshshell = WScript.CreateObject("WScript.Shell")

WshShell.RegWrite "HKCU\Software\Microsoft\Internet Explorer\Control Panel\", "Control Panel"


WshShell.RegWrite "HKCU\Software\Microsoft\Internet Explorer\Control Panel\AdvancedTab", "8080", "REG_DWORD"