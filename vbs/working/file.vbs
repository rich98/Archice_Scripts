Set objFSO = CreateObject("Scripting.FileSystemObject")
If objFSO.FileExists("C:\windows\cetova_install.txt") Then
    strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")
Set colSoftware = objWMIService.ExecQuery _
    ("SELECT * FROM Win32_Product WHERE Name = 'Cetova C-FAR Excel Add-In'")
For Each objSoftware in colSoftware
    objSoftware.Uninstall()
Next
Set objFSO = CreateObject("Scripting.FileSystemObject")
objFSO.DeleteFile("C:\windows\cetova_install.txt")
Wscript.Echo "cetova Excel plugin uninstalled please wait while new Cetova Excel plugin is installed"
Set WshShell = WScript.CreateObject ("WScript.Shell")

WshShell.Popup "installing new Excel plugin"

WshShell.Run "\\forum10\hotfix$\cetova2.exe /v/qn"

WScript.Sleep 300000

WshShell.Popup "finished installing new software Thank You"
Else
    Wscript.Echo "File does not exist.?"
End If
