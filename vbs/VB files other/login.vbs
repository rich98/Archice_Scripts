' login script for Forum Users

Set WshNetwork = CreateObject("WScript.Network")

On Error Resume Next

WshNetwork.RemoveNetworkDrive "t:" ,boolForce

WshNetwork.MapNetworkDrive "t:","\\forum10\share"

Set WshNetwork = CreateObject("WScript.Network")

PrinterPath = "\\forum-print\GENICOM COLOUR cl160"

WshNetwork.AddWindowsPrinterConnection PrinterPath

PrinterPath = "\\forum-print\HP 8500 Colour - PS"

WshNetwork.AddWindowsPrinterConnection PrinterPath

WScript.Sleep 3000

' Set WshShell = WScript.CreateObject ("WScript.Shell")

' WshShell.Run "t:\it\track-it\audit32.exe"



