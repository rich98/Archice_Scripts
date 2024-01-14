Dim fso, file
set fso = CreateObject ("Scripting.FileSystemObject")
set fso = fso.CreateFolder ("c:\filetest\subfolder")
set file = fso.CreateTextFile ("c:\filetest\filetest1.txt")