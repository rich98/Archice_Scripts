'folder_open.vbs - Fixes problem where Search opens on a double click


Set WshShell = WScript.CreateObject("WScript.Shell")

p1 = "HKEY_CLASSES_ROOT\Directory\shell\"
p2 = "none"

WshShell.RegWrite p1, p2

p1 = "HKEY_CLASSES_ROOT\Drive\shell\"
WshShell.RegWrite p1, p2

Set WshShell = Nothing

MyBox = MsgBox("Folders will now Open when double clicked", 4096, "Finished!")