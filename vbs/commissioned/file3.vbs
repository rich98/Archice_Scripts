Set objFSO = CreateObject("Scripting.FileSystemObject")
If objFSO.FileExists("C:\windows\cetova_install.txt") Then
    strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")
Set colSoftware = objWMIService.ExecQuery _
    ("SELECT * FROM Win32_Product WHERE Name = 'Cetova C-FAR Excel Add-In'")
For Each objSoftware in colSoftware
    objSoftware.Uninstall()
Set objFSO = CreateObject("Scripting.FileSystemObject")
objFSO.DeleteFile("C:\windows\cetova_install.txt")

Set WshShell = WScript.CreateObject ("WScript.Shell")

WshShell.Run "\\forum10\hotfix$\cetova2.exe /v/qn"

WScript.Sleep 3000
Next

Else
    Dim fso, file

set fso = CreateObject ("Scripting.FileSystemObject")

On Error Resume Next

set fso = CreateObject ("Scripting.FileSystemObject")

set file = fso.CreateTextFile ("c:\windows\cetova2.txt")

End if