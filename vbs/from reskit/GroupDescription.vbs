
'********************************************************************
'*
'* File:        GROUPDESCRIPTION.VBS
'* Created:     August 1998
'* Version:     1.0
'*
'* Main Function: Gets the description of user groups.
'* Usage: GROUPDESCRIPTION.VBS </A:adspath | /I:inputfile> [/O:outputfile]
'*        [/U:username] [/W:password] [/Q]
'*
'* Copyright (C) 1998 Microsoft Corporation
'*
'********************************************************************

OPTION EXPLICIT
ON ERROR RESUME NEXT

'Define constants
CONST CONST_ERROR                   = 0
CONST CONST_WSCRIPT                 = 1
CONST CONST_CSCRIPT                 = 2
CONST CONST_SHOW_USAGE              = 3
CONST CONST_PROCEED                 = 4

'Declare variables
Dim strADsPath, strInputFile, strOutputFile, strUserName, strPassword, intOpMode
Dim blnQuiet, i, strArgumentArray
ReDim strArgumentArray(0)

'Initialize variables
strArgumentArray(0) = ""
blnQuiet = False
strADsPath = ""
strInputFile = ""
strOutputFile = ""
strUserName = ""
strPassword = ""

'Get the command line arguments
For i = 0 to Wscript.arguments.count - 1
    ReDim Preserve strArgumentArray(i)
    strArgumentArray(i) = Wscript.arguments.item(i)
Next

'Check whether the script is run using CScript
Select Case intChkProgram()
    Case CONST_CSCRIPT
        'Do Nothing
    Case CONST_WSCRIPT
        WScript.Echo "Please run this script using CScript." & vbCRLF & _
            "This can be achieved by" & vbCRLF & _
            "1. Using ""CScript GROUPDESCRIPTION.vbs arguments"" for Windows 95/98 or" _
                & vbCRLF & _
            "2. Changing the default Windows Scripting Host setting to CScript" & vbCRLF & _
            "    using ""CScript //H:CScript //S"" and running the script using" & vbCRLF & _
            "    ""GROUPDESCRIPTION.vbs arguments"" for Windows NT."
        WScript.Quit
    Case Else
        WScript.Quit
End Select

'Parse the command line
intOpMode = intParseCmdLine(strArgumentArray, strADsPath, strInputFile, _
            strOutputFile, strUserName, strPassword, blnQuiet)
If Err.Number then
    Print " Error 0x" & CStr(Hex(Err.Number)) & " occurred in parsing the command line."
    If Err.Description <> "" Then
        Print "Error description: " & Err.Description & "."
    End If
    WScript.Quit
End If

Select Case intOpMode
    Case CONST_SHOW_USAGE
        Call ShowUsage()
    Case CONST_PROCEED
        Print " Working ... "
        Call GetDescription(strADsPath, strInputFile, strOutputFile, strUserName, strPassword)
    Case CONST_ERROR
        'Do nothing.
    Case Else                    'Default -- should never happen
        Print "Error occurred in passing parameters."
End Select

'********************************************************************
'*
'* Function intChkProgram()
'* Purpose: Determines which program is used to run this script.
'* Input:   None
'* Output:  intChkProgram is set to one of CONST_ERROR, CONST_WSCRIPT,
'*          and CONST_CSCRIPT.
'*
'********************************************************************

Private Function intChkProgram()

    ON ERROR RESUME NEXT

    Dim strFullName, strCommand, i, j

    'strFullName should be something like C:\WINDOWS\COMMAND\CSCRIPT.EXE
    strFullName = WScript.FullName
    If Err.Number then
        Print "Error 0x" & CStr(Hex(Err.Number)) & " occurred."
        If Err.Description <> "" Then
            Print "Error description: " & Err.Description & "."
        End If
        intChkProgram =  CONST_ERROR
        Exit Function
    End If

    i = InStr(1, strFullName, ".exe", 1)
    If i = 0 Then
        intChkProgram =  CONST_ERROR
        Exit Function
    Else
        j = InStrRev(strFullName, "\", i, 1)
        If j = 0 Then
            intChkProgram =  CONST_ERROR
            Exit Function
        Else
            strCommand = Mid(strFullName, j+1, i-j-1)
            Select Case LCase(strCommand)
                Case "cscript"
                    intChkProgram = CONST_CSCRIPT
                Case "wscript"
                    intChkProgram = CONST_WSCRIPT
                Case Else       'should never happen
                    Print "An unexpected program is used to run this script."
                    Print "Only CScript.Exe or WScript.Exe can be used to run this script."
                    intChkProgram = CONST_ERROR
            End Select
        End If
    End If

End Function

'********************************************************************
'*
'* Function intParseCmdLine()
'* Purpose: Parses the command line.
'* Input:   strArgumentArray    an array containing input from the command line
'* Output:  strADsPath          ADsPath of a group
'*          strInputFile        an input file name
'*          strOutputFile       an output file name
'*          strUserName         name of the current user
'*          strPassword         password of the current user
'*          blnQuiet            specifies whether to suppress messages
'*          intParseCmdLine     is set to one of CONST_ERROR, CONST_SHOW_USAGE, CONST_PROCEED.
'*
'********************************************************************

Private Function intParseCmdLine(strArgumentArray, strADsPath, strInputFile, _
                 strOutputFile, strUserName, strPassword, blnQuiet)

    ON ERROR RESUME NEXT

    Dim i, strFlag

    strFlag = strArgumentArray(0)

    If strFlag = "" then                'No arguments have been received
        Print "Arguments are required."
        intParseCmdLine = CONST_ERROR
        Exit Function
    End If

    If (strFlag="help") OR (strFlag="/h") OR (strFlag="\h") OR (strFlag="-h") _
        OR (strFlag = "\?") OR (strFlag = "/?") OR (strFlag = "?") OR (strFlag="h") Then
        intParseCmdLine = CONST_SHOW_USAGE
        Exit Function
    End If

    For i = 0 to UBound(strArgumentArray)
        strFlag = Left(strArgumentArray(i), InStr(1, strArgumentArray(i), ":")-1)
        If Err.Number Then            'An error occurs if there is no : in the string
            Err.Clear
            Select Case LCase(strArgumentArray(i))
                Case "/q"
                    blnQuiet = True
                Case else
                    Print "Invalid flag " & strArgumentArray(i) & "."
                    Print "Please check the input and try again."
                    intParseCmdLine = CONST_ERROR
                    Exit Function
            End Select
        Else
            Select Case LCase(strFlag)
                Case "/a"
                    strADsPath = FormatProvider(Trim(Right(strArgumentArray(i), Len(strArgumentArray(i))-3)))
                Case "/i"
                    strInputFile = Trim(Right(strArgumentArray(i), Len(strArgumentArray(i))-3))
                Case "/o"
                    strOutputFile = Trim(Right(strArgumentArray(i), Len(strArgumentArray(i))-3))
                Case "/u"
                    strUserName = Trim(Right(strArgumentArray(i), Len(strArgumentArray(i))-3))
                Case "/w"
                    strPassword = Trim(Right(strArgumentArray(i), Len(strArgumentArray(i))-3))
                Case else
                    Print "Invalid flag " & strFlag & "."
                    Print "Please check the input and try again."
                    intParseCmdLine = CONST_ERROR
                    Exit Function
            End Select
        End If
    Next

    'the root is required
    If strADsPath = "" and strInputFile = "" Then
        Print "Please enter ADsPath or InputFile."
        intParseCmdLine = CONST_ERROR
        Exit Function
    End If

    intParseCmdLine = CONST_PROCEED

End Function

'********************************************************************
'*
'* Sub ShowUsage()
'* Purpose: Shows the correct usage to the user.
'* Input:   None
'* Output:  Help messages are displayed on screen.
'*
'********************************************************************

Private Sub ShowUsage()

    Wscript.Echo ""
    Wscript.Echo "Gets the description of user groups." & vbCRLF
    Wscript.Echo "GROUPDESCRIPTION.VBS </A:adspath | /I:inputfile>"
    Wscript.Echo "[/U:username] [/W:password] [/Q]"
    Wscript.Echo "   /A, /I, /U, /W"
    Wscript.Echo "                 Parameter specifiers."
    Wscript.Echo "   adspath       The ADsPath of a user group."
    Wscript.Echo "   inputfile     A file containing ADsPaths of multiple user groups."
    Wscript.Echo "   username      Username of the current user."
    Wscript.Echo "   password      Password of the current user."
    Wscript.Echo "   /Q            Suppresses all output messages." & vbCRLF
    Wscript.Echo "EXAMPLE:"
    Wscript.Echo "GROUPDESCRIPTION.VBS /A:""WinNT://FooFoo/domain users"""
    Wscript.Echo "   prints the description of the ""domain users"" group of FooFoo."

End Sub

'********************************************************************
'*
'* Sub GetDescription()
'* Purpose: Gets the description of a group or groups.
'* Input:   strADsPath      the ADsPath of the group
'*          strInputFile    an input file name
'*          strOutputFile   an output file name
'*          strUserName     name of the current user
'*          strPassword     password of the current user
'* Output:  The description of the group is printed on the screen.
'*
'********************************************************************

Private Sub GetDescription(strADsPath, strInputFile, strOutputFile, strUserName, strPassword)

    ON ERROR RESUME NEXT

    Dim objFileSystem, objInputFile, objOutputFile

    If strInputFile <> "" or strOutputFile <> "" Then
    'Create a filesystem object
        set objFileSystem = CreateObject("Scripting.FileSystemObject")
        If Err.Number then
            Print "Error 0x" & CStr(Hex(Err.Number)) & " opening a filesystem object."
            If Err.Description <> "" Then
                Print "Error description: " & Err.Description & "."
            End If
            Exit Sub
        End If
    End If

    If strOutputFile <> "" Then
        'Open the file for output
        set objOutputFile = objFileSystem.OpenTextFile(strOutputFile, 8, True)
        If Err.Number then
            Print "Error 0x" & CStr(Hex(Err.Number)) & " opening file " & strOutputFile
            If Err.Description <> "" Then
                Print "Error description: " & Err.Description & "."
            End If
            Exit Sub
        End If
    Else
        objOutputFile = ""
    End If

    If strADsPath <> "" Then
        If blnGetOneDescription(strADsPath, strUserName, strPassword, objOutputFile) Then
            Exit Sub
        End If
    End If

    If strInputFile <> "" Then
        'Open the file for input
        set objInputFile = objFileSystem.OpenTextFile(strInputFile)
        If Err.Number then
            Print "Error 0x" & CStr(Hex(Err.Number)) & " opening file " & strInputFile
            If Err.Description <> "" Then
                Print "Error description: " & Err.Description & "."
            End If
            Exit Sub
        End If

        'Read from the input file
        While Not objInputFile.AtEndOfStream
            'Get rid of leading and trailing spaces
            strADsPath = Trim(objInputFile.ReadLine)
            If strADsPath <> "" Then
                Call blnGetOneDescription(strADsPath, strUserName, strPassword, objOutputFile)
            End If
        Wend
        objInputFile.Close
    End If

    If strOutputFile <> "" Then
        Print "Results are saved in file " & strOutputFile & "."
        objOutputFile.Close
    End If

End Sub


'********************************************************************
'*
'* Function blnGetOneDescription()
'* Purpose: Gets the description of a group.
'* Input:   strADsPath      the ADsPath of the group
'*          strInputFile    an input file name
'*          strOutputFile   an output file name
'*          strUserName     name of the current user
'*          strPassword     password of the current user
'* Output:  The description of the group is printed on the screen.
'*
'********************************************************************

Private Function blnGetOneDescription(strADsPath, strUserName, strPassword, objOutputFile)

    ON ERROR RESUME NEXT

    Dim objGroup, strProvider, objProvider, strDescription

    strDescription = ""
    blnGetOneDescription = False    'No error.

    If strUserName = ""    then        'The current user is assumed
        set objGroup = GetObject(strADsPath)
    Else                        'Credentials are passed
        strProvider = Left(strADsPath, InStr(1, strADsPath, ":"))
        set objProvider = GetObject(strProvider)
        'Use user authentication
        set objGroup = objProvider.OpenDsObject(strADsPath,strUserName,strPassword,1)
    End If
    If Err.Number then
		If CStr(Hex(Err.Number)) = "80070035" Then
			Print "Object " & strADsPath & " is not found."
		Else
			Print "Error 0x" & CStr(Hex(Err.Number)) & " occurred in getting object " _
				& strADsPath & "."
			If Err.Description <> "" Then
				Print "Error description: " & Err.Description & "."
			End If
		End If
		Err.Clear
        blnGetOneDescription = True        'An error occurred.
        Exit Function
    End If

    strDescription = objGroup.description
    If Err.Number Then
        Err.Clear
        blnGetOneDescription = True        'An error occurred.
        Exit Function
    End If

    If strDescription = "" Then
        WriteLine "There is no description for " & objGroup.ADsPath & ".", objOutputFile
    Else
        WriteLine """" & objGroup.ADsPath & """ description:", objOutputFile
        WriteLine "            " & strDescription, objOutputFile
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

'********************************************************************
'*
'* Sub Print()
'* Purpose: Prints a message on screen if blnQuiet = False.
'* Input:   strMessage      the string to print
'* Output:  strMessage is printed on screen if blnQuiet = False.
'*
'********************************************************************

Sub Print(ByRef strMessage)
    If Not blnQuiet then
        Wscript.Echo  strMessage
    End If
End Sub


'********************************************************************
'*
'* Function FormatProvider
'* Purpose: Formats Provider so it is not case sensitive
'* Input:   Provider    a string
'* Output:  FormatProvider is the Provider with the correct Case
'*
'********************************************************************

Private Function FormatProvider(Provider)
    FormatProvider = ""
    I = 1
    Do Until Mid(Provider, I, 1) = ":"
        If I = Len(Provider) Then
            'This Provider is Probabaly not valid, but we'll let it pass anyways.
            FormatProvider = Provider
            Exit Function
        End If
        I = I + 1
    Loop

    Select Case LCase(Left(Provider, I - 1))
        Case "winnt"
            FormatProvider = "WinNT" & Right(Provider,Len(Provider) - (I - 1))
        Case "ldap"
            FormatProvider = "LDAP" & Right(Provider,Len(Provider) - (I - 1))			
    End Select


End Function


'********************************************************************
'*                                                                  *
'*                           End of File                            *
'*                                                                  *
'********************************************************************

'********************************************************************
'*
'* Procedures calling sequence: GROUPDESCRIPTION.VBS
'*
'*  intChkProgram
'*  intParseCmdLine
'*  ShowUsage
'*  GetDescription
'*      blnGetOneDescription
'*          WriteLine
'*
'********************************************************************
