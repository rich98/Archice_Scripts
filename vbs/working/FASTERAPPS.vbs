set Wshshell = WScript.CreateObject("WScript.Shell")

WshShell.RegWrite "HKLM\SYSTEM\CurrentControlSet\Control\Session Manager\Memory Management\PrefetchParameters\EnablePrefetcher", "30", "REG_DWORD"