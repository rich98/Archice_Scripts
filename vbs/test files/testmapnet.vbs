Set WshNetwork = CreateObject("WScript.Network")

WshNetwork.MapNetworkDrive "x:","\\forum10\home\" & UserName "

