strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")
Set colFolders = objWMIService.ExecQuery _
    ("SELECT * FROM Win32_Directory WHERE Name = 'c:\\moc'")
For Each objFolder in colFolders
    
Next
