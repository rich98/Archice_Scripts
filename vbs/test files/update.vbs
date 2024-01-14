set Wshshell = WScript.CreateObject("WScript.Shell")

On Error Resume Next

strKeyValue = "HKLM\CLIENTHOTFIX\Q1234"

return = WshShell.RegRead strKeyValue

Set WshNetwork = CreateObject("WScript.Network")

WshNetwork.MapNetworkDrive "u:","\\matrix2000\hotfix$"

Set WshShell = WScript.CreateObject ("WScript.Shell")

WshShell.Run "u:\ update name"

WshShell.RegWrite "HKLM\CLIENTHOTFIX\Q1234"

