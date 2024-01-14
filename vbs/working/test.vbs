set Wshshell = WScript.CreateObject("WScript.Shell")

WshShell.RegWrite "HKCU\Software\Microsoft\Windows\CurrentVersion\Internet Settings\5.0\Cache\Content\CacheLimit", "1024", "REG_DWORD"