Set WshNetwork = CreateObject("WScript.Network")
strSharePath = "\\forum10\home"
strName = "UserName"
WshNetwork.MapNetworkDrive "J:" strSharePath _, strName



