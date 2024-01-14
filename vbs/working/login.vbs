' login script for Forum Users

Set WshNetwork = CreateObject("WScript.Network")

On Error Resume Next

WshNetwork.RemoveNetworkDrive "t:" ,boolForce

WshNetwork.MapNetworkDrive "t:","\\forum10\share"

WScript.Sleep 3000

' Set WshShell = WScript.CreateObject ("WScript.Shell")

' WshShell.Run "t:\it\track-it\audit32.exe"



