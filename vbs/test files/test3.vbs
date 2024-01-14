Set WshNetwork = CreateObject("WScript.Network")

strUserName = WshNetwork.UserName

strDrive = "q:"

strShare = "\\forum10\home\ " & strUserName & "

WshNetwork.MapNetworkDrive strDrive, strShare 