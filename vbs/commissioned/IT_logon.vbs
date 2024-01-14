' IT Login Script 

Set WshNetwork = CreateObject("WScript.Network")

On Error Resume Next

WshNetwork.RemoveNetworkDrive "p:" ,boolForce

On Error Resume Next

WshNetwork.RemoveNetworkDrive "t:" ,boolForce

WshNetwork.MapNetworkDrive "t:","\\forum10\share"

WshNetwork.MapNetworkDrive "p:","\\matrix2000\apps"

Wscript.Sleep 3000

' ***** start apps faster *****

Set WshShell = WScript.CreateObject ("WScript.Shell")

WshShell.RegWrite "HKLM\SYSTEM\CurrentControlSet\Control\Session Manager\Memory Management\PrefetchParameters\EnablePrefetcher", "30", "REG_DWORD"

WshShell.Run "t:\it\track-it\audit32.exe"



