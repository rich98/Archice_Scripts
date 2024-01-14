Set WshShell = WScript.CreateObject ("WScript.Shell")

WshShell.Run "net stop spooler"

wscript.sleep 6000

WshShell.Run "net start spooler"