Set WshShell = WScript.CreateObject ("WScript.Shell")

WshShell.Popup "Running Machine Audit"

WshShell.Run "t:\it\track-it 6\audit32.exe"

WScript.Sleep 300000

WshShell.Popup "finished Running Audit thank you"
