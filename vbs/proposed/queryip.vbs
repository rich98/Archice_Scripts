strComputer = "."
Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
Set IPConfigSet = objWMIService.ExecQuery _
("Select IPAddress from Win32_NetworkAdapterConfiguration where IPEnabled=TRUE")
For Each IPConfig in IPConfigSet
If Not IsNull(IPConfig.IPAddress) Then 
For i=LBound(IPConfig.IPAddress) to UBound(IPConfig.IPAddress)
WScript.Echo IPConfig.IPAddress(i)
Next
End If
Next




