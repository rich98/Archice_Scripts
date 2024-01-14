Dim fso, file1

set fso = CreateObject ("Scripting.FileSystemObject")

On Error Resume Next

set fso = fso.CreateFolder ("c:\filetest")

WScript.sleep 10000

set fso = CreateObject ("Scripting.FileSystemObject")

set file1 = fso.CreateTextFile ("c:\filetest\testing1.txt")

set file1 = fso.CreateTextFile ("c:\filetest\testing2.txt")

