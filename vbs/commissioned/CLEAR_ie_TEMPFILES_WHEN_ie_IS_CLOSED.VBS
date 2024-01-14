Set WshShell = WScript.CreateObject ("WScript.Shell")


WshShell.Regwrite "HKCU\Software\Microsoft\Windows\CurrentVersion\Internet Settings\Cache\Persistent", "0", "REG_DWORD"