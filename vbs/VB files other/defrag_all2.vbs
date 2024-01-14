'defrag_all2.vbs
'Defrags all hard disks - Can be run as a Scheduled Task
'Modified to create an error log and display it
'© Doug Knox - 4/13/2002
'This code may be freely distributed/modified

Option Explicit

Dim WshShell, fso, d, dc, ErrStr(), Return, X, A(), MyFile, I, MyBox, Drive

Set WshShell = WScript.CreateObject("WScript.Shell")
Set fso = CreateObject("Scripting.FileSystemObject")
X = 0

   Set dc = fso.Drives
For Each d in DC 
	If d.DriveType = 2 Then
	   X = X + 1

'Determine drive letter of first fixed disk
'This is the drive that the error report will be placed on
		If X = 1 Then
		   Drive = d
		End If
	End If
Next

ReDim A(X)
ReDim ErrStr(X)

X = 0
For Each d in dc
      If d.DriveType = 2 Then
      X = X + 1
      Return = WshShell.Run("defrag " & d & " -f", 1, TRUE)

'Determine the Error code returned by Defrag for the current drive and save it
If return = 0 then
  ErrStr(x) = ErrStr(x) &  "Drive " & d & " Defrag completed successfully" & vbCRLF
elseif return = 1 then
  ErrStr(x) = ErrStr(x) &  "Drive " & d & " Defrag aborted with error level " & return & " (defrag was cancelled manually) " & vbCRLF
elseif return = 2 then
  ErrStr(x) = ErrStr(x) &  "Drive " & d & " Defrag aborted with error level " & return & " (there was a command line error. Check your command line for valid switches and drives)" & vbCRLF
elseif return = 3 then
  ErrStr(x) = ErrStr(x) &  "Drive " & d & " Defrag aborted with error level " & return & " (there was an unknown error)" & vbCRLF
elseif return = 4 then
  ErrStr(x) = ErrStr(x) &  "Drive " & d & " Defrag aborted with error level " & return & " (defrag could not run due to insufficient memory resources)" & vbCRLF
  'errorlevel 5 is not currently used
elseif return = 5 then
  ErrStr(x) = ErrStr(x) &  "Drive " & d & " Defrag aborted with error level " & return & " (general error)" & vbCRLF
elseif return = 6 then
  ErrStr(x) = ErrStr(x) &  "Drive " & d & " Defrag aborted with error level " & return & "(System error: either the account used to run defrag is not an administrator, there is a problem loading the resource DLL, or a defrag engine could not be accessed. Check for proper user permissions and run Sfc.exe to validate system files)" & vbCRLF
elseif return = 7 then
  ErrStr(x) = ErrStr(x) &  "Drive " & d & " Defrag aborted with error level " & return & " (There is not enough free space on the drive. Defrag needs 15% free space to run on a volume)" & vbCRLF
else
  ErrStr(x) = ErrStr(x) &  "Drive " & d & " Defrag aborted with an unknown error level: " & return & vbCRLF
end if

       End If
   Next

'Create the Error Report in the root of the first fixed disk.
Set MyFile = fso.OpenTextFile(Drive & "\defragreport.txt", 2, True)
MyFile.WriteLine(Date) & vbCRLF
MyFile.WriteLine(Time) & vbCRLF
   For I = 1 to X
      MyFile.WriteLine(ErrStr(I))
   Next
   MyFile.Close

Return = WshShell.Run(Drive & "\defragreport.txt",3,True)

Set WshShell = Nothing
Set fso = Nothing