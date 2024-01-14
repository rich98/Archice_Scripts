'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' 
'  File:           DeviceMem.VBS
'  Created:        December 1998
'  Version:        1.0
' 
'  Main Function: Outputs Information on Disk Drives.
'
'  Drives.vbs  [/S <server>] [/O <outputfile>] [/U <username>] [/W <password>] 
' 
'  Copyright (C) 1998 Microsoft Corporation
' 
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
CONST THISFILE                      = "DeviceMem.VBS"


'Define constants
CONST CONST_ERROR                   = 0
CONST CONST_WSCRIPT                 = 1
CONST CONST_CSCRIPT                 = 2
CONST CONST_SHOW_USAGE              = 3
CONST CONST_PROCEED                 = 4
CONST CONST_LIST                    = "LIST"
CONST CONST_DEFAULTTASK             = "LIST"

'Generic variables for remote automation
Dim intOpMode,    strTaskCommand,   strArgArray(),   i 'generic loop counter
Dim strServer,    strUserName,      strPassword,     objService
Dim blnReboot,    blnForce,         blnVerbose


'Get the command line arguments
ReDim strArgArray (Max (0, Wscript.arguments.count - 1))

For i = 0 to Wscript.arguments.count - 1
    strArgArray(i) = Wscript.arguments.Item(i)
Next

call VerifyCScript()

'Parse the command line
intOpMode = intParseCmdLine(  strArgArray,      _
                              strTaskCommand,   _ 
                              strServer,        _
                              strUserName,      _
                              strPassword,      _
                              strOutPutFile    )

Do
  Select Case intOpMode
    Case CONST_PROCEED    Exit Do
    Case CONST_SHOW_USAGE Call ShowUsage()
    Case Else             Call Print("Error occurred in passing parameters.")
  End Select
  Wscript.Quit
Loop

'Establish a connection with the server
Call blnConnect (objService, strServer, "root/cimv2", strUserName, strPassword)

if not isobject (objService) then 
  call Print("Please check the server name, credentials and WBEM Core.")
  Wscript.Quit
end if

If IsEmpty(strServer) Then
  Dim objWshNet
  Set objWshNet = CreateObject("Wscript.Network")
  strServer = objWshNet.ComputerName
End If

Select Case strTaskCommand

    Case CONST_LIST
		call List (objService)

	Case Else
		call ShowUsage ()

End Select

'END of main routine


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' 
'  Function List ()
'  Purpose: List instances of the class Win32_DeviceMemoryAddress
' 
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub List (objService)

  ON ERROR RESUME NEXT

  Dim objFileSystem, objOutFile
  Dim strClass, objSet, obj, objInst, strLine

  If NOT IsEmpty(strOutputFile) Then
    'Create a file object
    set objFileSystem = CreateObject("Scripting.FileSystemObject")
    If ErrorOccurred("opening a filesystem object.") Then Exit Sub
    'Open the file for output
    set objOutFile = objFileSystem.OpenTextFile(strOutputFile, 8, True)
    If ErrorOccurred("opening file " + strOutputFile) Then Exit Sub
  End If

  'Get the first instance
  strClass = "Win32_DeviceMemoryAddress"
  Set objSet = objService.InstancesOf(strClass)
  If ErrorOccurred ("obtaining the "& strClass) Then Exit Sub

  WriteLine "Instances of "& strClass& " on "& strServer, objOutFile
  Dim intW(5)
  intW(0) = 23
  intW(1) = 11
  intW(2) = 12
  intW(3) = 14
  intW(4) = 14
  strLine = Empty
  strLine = strLine + strPackString ("Range"       ,intW(0),1,1)
  strLine = strLine + strPackString ("Starting"    ,intW(1),1,1)
  strLine = strLine + strPackString ("Ending"      ,intW(2),1,1)
  strLine = strLine + strPackString ("MemoryType"  ,intW(3),1,1)
  strLine = strLine + strPackString ("Status"      ,intW(4),1,1)
  WriteLine " ", objOutFile
  WriteLine strLine, objOutFile
  WriteLine Replace (Space(73), " ", "-"), objOutFile

  For Each obj In objSet
    strLine = Empty
    strLine = strLine + strPackString (obj.Caption        ,intW(0),1,1) 
    strLine = strLine + strPackString (obj.StartingAddress,intW(1),1,1)
    strLine = strLine + strPackString (obj.EndingAddress  ,intW(2),1,1)
    strLine = strLine + strPackString (obj.MemoryType     ,intW(3),1,1)
    strLine = strLine + strPackString (obj.Status         ,intW(4),1,1)
    WriteLine strLine, objOutFile
  Next

  If NOT IsEmpty(objOutFile) Then
    objOutFile.Close
    Wscript.Echo "Results are saved in file " & strOutputFile & "."
  End If

End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' 
'  Function blnConnect()
'  Purpose: Connects to machine strServer.
'  Input:   strServer       a machine name
'           strNameSpace    a namespace
'           strUserName     name of the current user
'           strPassword     password of the current user
'  Output:  objService is returned  as a service object.
' 
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function blnConnect (objService, strServer, strNameSpace, strUserName, strPassword)

    ON ERROR RESUME NEXT
    Dim objLocator

    blnConnect = True 'Return an error if we exit before reaching the end.

    'Create Locator object to connect to remote CIM object manager
    Set objLocator = CreateObject("WbemScripting.SWbemLocator")
		if ErrorOccurred("occurred creating a locator object.") _
		Then Exit Function

    'Connect to the namespace which is either local or remote
    Set objService = objLocator.ConnectServer (strServer, strNameSpace, _
                                               strUserName, strPassword)
 		if ErrorOccurred("occurred connecting to server \\" & strServer & "." ) _
		then Exit Function

		ObjService.Security_.impersonationlevel = 3
 		if ErrorOccurred("occurred setting security level on \\" & strServer & "." ) _
		then Exit Function

    blnConnect = False     'There is no error.

End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' 
'  Function ErrorOccurred()
'  Purpose: Reports error with a string saying what the error occurred in.
'  Input:   strIn		string saying what the error occurred in.
'  Output:				displayed on screen 
' 
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function ErrorOccurred (strIn)
    If Err.Number then
        call Print( "Error 0x" & CStr(Hex(Err.Number)) & " occurred " & strIn)
        If Err.Description <> "" Then
            call Print( "Error description: " & Err.Description)
        End If
        Err.Clear
		ErrorOccurred = true
    Else
		ErrorOccurred = false
    End If
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' 
'  Sub Print()
'  Purpose: Prints a message on screen.
'  Input:   strMessage      the string to print
'  Output:  strMessage is printed on screen.
' 
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub Print (ByRef strMessage)

    Wscript.Echo  strMessage

End Sub

'********************************************************************
'*
'* Function strPackString()
'* Purpose: Attaches spaces to a string to increase the length to intWidth.
'* Input:   strString   a string
'*          intWidth   the intended length of the string
'*          blnAfter    specifies whether to add spaces after or before the string
'*          blnTruncate specifies whether to truncate the string or not if
'*                      the string length is longer than intWidth
'* Output:  strPackString is returned as the packed string.
'*
'********************************************************************
Private Function strPackString(strString, ByVal intWidth, blnAfter, blnTruncate)

    'ON ERROR RESUME NEXT

    intWidth = CInt(intWidth)
    blnAfter = CBool(blnAfter)
    blnTruncate = CBool(blnTruncate)
    If Err.Number Then
        Print "Argument type is incorrect!"
        Err.Clear
        Wscript.Quit
    End If

    If IsNull(strString) Then
        strPackString = "null" & Space(intWidth-4)
        Exit Function
    End If

    strString = CStr(strString)
    If Err.Number Then
        Print "Argument type is incorrect!"
        Err.Clear
        Wscript.Quit
    End If

    If intWidth > Len(strString) Then
        If blnAfter Then
            strPackString = strString & Space(intWidth-Len(strString))
        Else
            strPackString = Space(intWidth-Len(strString)) & strString & " "
        End If
    Else
        If blnTruncate Then
            strPackString = Left(strString, intWidth-1) & " "
        Else
            strPackString = strString & " "
        End If
    End If

End Function

'********************************************************************
'*
'* Sub WriteLine()
'* Purpose: Writes a text line either to a file or on screen.
'* Input:   strMessage  the string to print
'*          objFile     an output file object
'* Output:  strMessage is either displayed on screen or written to a file.
'*
'********************************************************************
Sub WriteLine(ByRef strMessage, ByRef objFile)

    If IsObject(objFile) then        'objFile should be a file object
        objFile.WriteLine strMessage
    Else
        Wscript.Echo  strMessage
    End If

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' 
'  Sub Max()
' 
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function Max (intA, intB) 
    Max = intA
    if intA < intB then Max = intB
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
'  Sub VerifyCScript ()
'
'  Purpose:  Verify that the script is called with CScript, and 
'            take appropriate action (quit with GUI usage message)
'
'  Suggested Use: add this line at beginning of main module:
'            call VerifyCScript()
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub VerifyCScript ()

  Select Case intChkProgram()
    Case CONST_CSCRIPT 
      'Do Nothing
    Case CONST_WSCRIPT
      WScript.Echo "Please run this script using CScript." & vbCRLF & _
        "This can be achieved by" & vbCRLF & _
        "1. Using ""CScript "&THISFILE&" arguments"" for Windows 95/98 or" & vbCRLF & _
        "2. Changing the default Windows Scripting Host setting to CScript" & vbCRLF & _
        "    using ""CScript //H:CScript //S"" and running the script using" & vbCRLF & _
        "    """&THISFILE&" arguments"" for Windows NT."
      WScript.Quit
    Case Else
      WScript.Quit
  End Select

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
'  Function intChkProgram()
'
'  Purpose: Determines which program is used to run this script.
' 
'  Input:   None
'
'  Output:  intChkProgram is set to one of CONST_ERROR, CONST_WSCRIPT,
'           and CONST_CSCRIPT.
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function intChkProgram()

    ON ERROR RESUME NEXT

    Dim strFullName, strCommand, i, j

    intChkProgram =  CONST_ERROR

    'strFullName should be something like C:\WINDOWS\COMMAND\CSCRIPT.EXE
    strFullName = WScript.FullName
    If ErrorOccurred("occurred.") Then Exit Function

    i = InStr(1, strFullName, ".exe", 1)
    If i = 0 Then Exit Function

    j = InStrRev(strFullName, "\", i, 1)
    If j = 0 Then Exit Function

    strCommand = Mid(strFullName, j+1, i-j-1)
    Select Case LCase(strCommand)
        Case "cscript"
            intChkProgram = CONST_CSCRIPT
        Case "wscript"
            intChkProgram = CONST_WSCRIPT
        Case Else       'should never happen
            call Print( "An unexpected program is used to run this script.")
            call Print( "Only CScript.Exe or WScript.Exe can be used to run this script.")
    End Select

End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' 
'  Function GetArg()
'
'  Purpose: Helper to intParseCmdLine()
' 
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function GetArg (strArgArray, intArgIter, strVar, StrVarName) 

  GetArg = True 'return error if we don't make it to end function

  intArgIter = intArgIter + 1
  If intArgIter > UBound(strArgArray) Then Exit Function

  strVar = strArgArray (intArgIter)
  If Err.Number Then
    call Print( "Invalid " & StrVarName & ".")
    call Print( "Please check the input and try again.")
  Exit Function
  End If

  If InStr(strVar, "/") Then
    call Print( "Invalid " & StrVarName)
    call Print( "Please check the input and try again.")
    Wscript.Quit
  End If

  GetArg = False 'success

End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'  Function ShowUsage ()
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub ShowUsage ()

  Wscript.Echo ""
  Wscript.Echo "Outputs Information on Device Memory Address Ranges."
  Wscript.Echo ""
  Wscript.Echo "SYNTAX:"
  Wscript.Echo "  DeviceMem.vbs  [/S <server>] [/O <outputfile>] [/U <username>]"
  Wscript.Echo "  [/W <password>]"
  Wscript.Echo ""
  Wscript.Echo "PARAMATER SPECIFIERS:"
  Wscript.Echo "   server        A machine name."
  Wscript.Echo "   outputfile    The output file name."
  Wscript.Echo "   username      The current user's name."
  Wscript.Echo "   password      Password of the current user."
  Wscript.Echo ""
  Wscript.Echo "EXAMPLE:"
  Wscript.Echo "1. cscript DeviceMem.vbs"
  Wscript.Echo "   Get the device memory address ranges for the current machine."
  Wscript.Echo "2. cscript DeviceMem.vbs /S MyMachine2"
  Wscript.Echo "   Get the device memory address ranges  for the machine MyMachine2."

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' 
'  Function intParseCmdLine()
'  Purpose: Parses the command line.
'  Input:   strArgArray  an array containing input from the command line
'  Output:  strTaskCommand    one of help, ...
'           strServer         a machine name
'           strUserName       the current user's name
'           strPassword       the current user's password
'           strOutPutFile     an output file
'           intParseCmdLine   is set to one of CONST_ERROR, 
'                             CONST_SHOW_USAGE, or CONST_PROCEED.
' 
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function intParseCmdLine( strArgArray,      _
                                  strTaskCommand,   _ 
                                  strServer,        _
                                  strUserName,      _
                                  strPassword,      _
                                  strOutPutFile    )

  Dim strFlag
  Dim intArgIter
  Dim objFileSystem

  strFlag = strArgArray(0)
  If strFlag = "" Then                 'No arguments have been received
    strTaskCommand = CONST_DEFAULTTASK 'Default Task
    intParseCmdLine = CONST_PROCEED
    Exit Function
  End If

  'Check if the user is asking for help or is just confused
  intParseCmdLine = CONST_SHOW_USAGE
  If (strFlag="help") OR (strFlag="/h") OR (strFlag="\h") OR (strFlag="-h") _
   OR (strFlag= "\?") OR (strFlag="/?") OR (strFlag="?") OR (strFlag="h") _
  Then Exit Function

  'Retrieve the command line and set appropriate variables
  Dim intArgMax
  intArgMax = UBound(strArgArray)

  'Return an error if we don't survive the next round
  intParseCmdLine = CONST_ERROR 

  intArgIter = 0
  Do While intArgIter <= UBound(strArgArray)
    Select Case LCase(strArgArray(intArgIter))

      Case "/s"
        If GetArg (strArgArray, intArgIter, strServer, "server name") _
        Then Exit Function

      Case "/u" 
        If GetArg (strArgArray, intArgIter, strUserName, "user name") _
        Then Exit Function

      Case "/w"
        If GetArg (strArgArray, intArgIter, strPassword, "password") _
        Then Exit Function

      Case "/o"
        If GetArg (strArgArray, intArgIter, strOutPutFile, "output filename") _
        Then Exit Function

      Case Else 'We shouldn't get here
        call Print( "Invalid or misplaced parameter: " _
                  & strArgArray(intArgIter) & vbCRLF _
                  & "Please check the input and try again," & vbCRLF _
                  & "or invoke with '/?' for help with the syntax." )
        Wscript.Quit

    End Select

    intArgIter = intArgIter + 1     'default advance

  Loop '** intArgIter <= UBound(strArgArray) **

  'we survived the last round; return default command if not set
  If IsEmpty(strTaskCommand) _
    Then strTaskCommand = CONST_DEFAULTTASK 'Default Task

  intParseCmdLine = CONST_PROCEED

End Function
