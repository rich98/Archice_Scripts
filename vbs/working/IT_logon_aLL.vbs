' IT Login Script 

Set WshNetwork = CreateObject("WScript.Network")

On Error Resume Next

WshNetwork.RemoveNetworkDrive "p:" ,boolForce

On Error Resume Next

WshNetwork.RemoveNetworkDrive "t:" ,boolForce

WshNetwork.MapNetworkDrive "t:","\\forum10\share"

WshNetwork.MapNetworkDrive "p:","\\matrix2000\apps"

Wscript.Sleep 3000

' Set WshShell = WScript.CreateObject ("WScript.Shell")

' WshShell.Run "t:\it\track-it\audit32.exe"

set Wshshell = WScript.CreateObject("WScript.Shell")

' set proxy and port values

WshShell.Regwrite "HKLM\SOFTWARE\Network Associates\TVD\Shared Components\McUpdate\CurrentVersion\Update\Update Site1\dwProxyPort", "8080", "REG_DWORD"

WshShell.Regwrite "HKLM\SOFTWARE\Network Associates\TVD\Shared Components\McUpdate\CurrentVersion\Update\Update Site1\szProxy", "VPN1"

' Remove proxy and port values

'WshShell.Regwrite "HKLM\SOFTWARE\Network Associates\TVD\Shared Components\McUpdate\CurrentVersion\Update\Update Site1\dwProxyPort", ""


'WshShell.Regwrite "HKLM\SOFTWARE\Network Associates\TVD\Shared Components\McUpdate\CurrentVersion\Update\Update Site1\szProxy", "" 

