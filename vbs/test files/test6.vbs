Set WshNetwork = CreateObject("WScript.Network")

On Error Resume Next

WshNetwork.RemoveNetworkDrive "t:" ,boolForce

Wscript.sleep 10000

Dim dtmValue

dtmValue = Now

MsgBox  Now
