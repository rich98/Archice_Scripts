On Error Resume Next
strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")
Set colSoftware = objWMIService.ExecQuery _
    ("SELECT * FROM Win32_Product WHERE Name = 'Cetova C-FAR Excel Add-In'")
For Each objSoftware in colSoftware
    objSoftware.Uninstall()
Next



