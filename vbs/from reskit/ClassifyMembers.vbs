
'********************************************************************
'*
'* File:        CLASSIFYMEMBERS.VBS
'* Created:     August 1998
'* Version:     1.0
'*
'* Main Function: Lists all Members of a container or a group.
'* Usage: CLASSIFYMEMBERS.VBS adspath [/O:outputfile] [/U:username] [/W:password] [/Q]
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
Dim strADsPath, strUserName, strPassword, strOutputFile
Dim i, intOpMode
Redim strArgumentArray(0)

'Initialize variables
strArgumentArray(0) = ""
strADsPath = ""
strUserName = ""
strPassword = ""
strOutputFile = ""

'Get the command line arguments
For i = 0 to Wscript.arguments.count - 1
    Redim Preserve strArgumentArray(i)
    strArgumentArray(i) = Wscript.arguments.item(i)
Next

'Check whether the script is run using CScript
Select Case intChkProgram()
    Case CONST_CSCRIPT
        'Do Nothing
    Case CONST_WSCRIPT
        WScript.Echo "Please run this script using CScript." & vbCRLF & _
            "This can be achieved by" & vbCRLF & _
            "1. Using ""CScript CLASSIFYMEMBERS.vbs arguments"" for Windows 95/98 or" _
                & vbCRLF & _
            "2. Changing the default Windows Scripting Host setting to CScript" & vbCRLF & _
            "    using ""CScript //H:CScript //S"" and running the script using" & vbCRLF & _
            "    ""CLASSIFYMEMBERS.vbs arguments"" for Windows NT."
        WScript.Quit
    Case Else
        WScript.Quit
End Select

'Parse the command line
intOpMode = intParseCmdLine(strArgumentArray, strADsPath, _
            strOutputFile, strUserName, strPassword)
If Err.Number then
    Print "Error 0x" & CStr(Hex(Err.Number)) & " occurred in parsing the command line."
    If Err.Description <> "" Then
        Print "Error description: " & Err.Description & "."
    End If
    WScript.Quit
End If

Select Case intOpMode
    Case CONST_SHOW_USAGE
        Call ShowUsage()
    Case CONST_PROCEED
        Call GetMembers(strADsPath, strUserName, strPassword, strOutputFile)
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
'* Output:  strADsPath          ADsPath of a group or container object
'*          strUserName         name of the current user
'*          strPassword         password of the current user
'*          strOutputFile       an output file name
'*          intParseCmdLine     is set to one of CONST_ERROR, CONST_SHOW_USAGE, CONST_PROCEED.
'*
'********************************************************************

Private Function intParseCmdLine(strArgumentArray, strADsPath, _
    strOutputFile, strUserName, strPassword)

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

    strADsPath = FormatProvider(strFlag)            'The first parameter must be ADsPath of the object.

    For i = 1 to UBound(strArgumentArray)
        strFlag = Left(strArgumentArray(i), InStr(1, strArgumentArray(i), ":")-1)
        If Err.Number Then            'An error occurs if there is no : in the string
            Err.Clear
            Select Case LCase(strArgumentArray(i))
                Case else
                    Print "Invalid flag " & strArgumentArray(i) & "."
                    Print "Please check the input and try again."
                    intParseCmdLine = CONST_ERROR
                    Exit Function
            End Select
        Else
            Select Case LCase(strFlag)
                Case "/o"
                    strOutputFile = Right(strArgumentArray(i), Len(strArgumentArray(i))-3)
                Case "/u"
                    strUserName = Right(strArgumentArray(i), Len(strArgumentArray(i))-3)
                Case "/w"
                    strPassword = Right(strArgumentArray(i), Len(strArgumentArray(i))-3)
                Case else
                    Print "Invalid flag " & strFlag & "."
                    Print "Please check the input and try again."
                    intParseCmdLine = CONST_ERROR
                    Exit Function
            End Select
        End If
    Next

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

    Wscript.echo ""
    Wscript.echo "Lists all members of a container or group object. In case of a"
    Wscript.echo "container, the member objects are grouped according to the class."
    Wscript.echo ""
    Wscript.echo "CLASSIFYMEMBERS.VBS adspath [/U:username] [/W:password] [/O:outputfile]"
    Wscript.echo ""
    Wscript.echo "Parameter specifiers:"
    Wscript.echo "   adspath       ADsPath of a container or group object."
    Wscript.echo "   username      Username of the current user."
    Wscript.echo "   password      Password of the current user."
    Wscript.Echo "   outputfile    The output file name."
    Wscript.echo ""
    Wscript.Echo "EXAMPLE:"
    Wscript.echo "CLASSIFYMEMBERS.VBS WinNT://FooFoo"
    Wscript.echo "   lists all members of FooFoo with the result sorted"
    Wscript.echo "   according to the class type."

End Sub

'********************************************************************
'*
'* Sub GetMembers()
'* Purpose: Lists all members of a container or group object.
'* Input:   strADsPath      ADsPath of a group or container object
'*          strUserName     name of the current user
'*          strPassword     password of the current user
'*          strOutputFile   an output file name
'* Output:  ADsPaths of the member objects are either printed on screen or saved
'*          in strOutputFile. The ADsPaths are sorted according to the class type.
'*
'********************************************************************

Private Sub GetMembers(strADsPath, strUserName, strPassword, strOutputFile)

    ON ERROR RESUME NEXT

    Dim strProvider, objProvider, objADs, objFileSystem, objOutputFile
    Dim objSchema, strClassArray(), objMember, i, intCount
    Redim strClassArray(0)

    strClassArray(0) = ""
	intCount = 0

    Print "Getting object " & strADsPath & "..."
    If strUserName = ""    then        'The current user is assumed
        set objADs = GetObject(strADsPath)
    Else                        'Credentials are passed
        strProvider = Left(strADsPath, InStr(1, strADsPath, ":"))
        set objProvider = GetObject(strProvider)
        'Use user authentication
        set objADs = objProvider.OpenDsObject(strADsPath,strUserName,strPassword,1)
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
        Exit Sub
    End If

    If strOutputFile = "" Then
        objOutputFile = ""
    Else
        'After discovering the object, open a file to save the results
        'Create a filesystem object
        set objFileSystem = CreateObject("Scripting.FileSystemObject")
        If Err.Number then
            Print "Error 0x" & CStr(Hex(Err.Number)) & " opening a filesystem object."
            If Err.Description <> "" Then
                Print "Error description: " & Err.Description & "."
            End If
            objOutputFile = ""
        Else
            'Open the file for output
            set objOutputFile = objFileSystem.OpenTextFile(strOutputFile, 8, True)
            If Err.Number then
                Print "Error 0x" & CStr(Hex(Err.Number)) & " opening file " & strOutputFile
                If Err.Description <> "" Then
                    Print "Error description: " & Err.Description & "."
                End If
                objOutputFile = ""
            End If
        End If
    End If

    'Get the object that holds the schema
    If strUserName = ""    then                                'The current user is assumed
        set objSchema = GetObject(objADs.schema)
    Else
    'Use user authentication
        set objSchema = objProvider.OpenDsObject(objADs.schema,strUserName,strPassword,1)
    End If
    If Err.Number then                'Can not get the schema for this object
        Print "Error 0x" & CStr(Hex(Err.Number)) & " occurred in getting the object schema."
        If Err.Description <> "" Then
            Print "Error description: " & Err.Description & "."
        End If
        Err.Clear
        'Now just list the members directly
        intCount = intGetDirectMembers(objADs, "", "", objOutputFile)
    Else                                            'The schema is found
        If objSchema.Container then            'If it is a container object
            'Let's determine which classes can be contained in this object
            i = 0
            For Each objMember in objSchema.Containment
                Redim Preserve strClassArray(i)
                strClassArray(i) = CStr(objMember)
                i = i + 1
            Next
            If strClassArray(0) = "" Then        'Nothing is found in this container's schema.
                intCount = intGetDirectMembers(objADs, objSchema, "", objOutputFile)
            Else
                For i =0 to UBound(strClassArray)
                    intCount = intCount + intGetDirectMembers(objADs, objSchema, _
                        strClassArray(i), objOutputFile)
                next
            End If
        Else    'It is a leaf object. Only group members are possible
            intCount = intGetDirectMembers(objADs, objSchema, "", objOutputFile)
        End If
    End If

    If intCount = 0 then                        'Nothing has been found
        Print "Object " & objADs.ADsPath & """ does not have any members."
    Else
        Print "Object " & objADs.ADsPath & """ has " & intCount & " members."
    End If

    If strOutputFile <> "" Then
        Wscript.echo "Results are saved in file " & strOutputFile & "."
        objOutputFile.Close
    End If

End Sub

'********************************************************************
'*
'* Function intGetDirectMembers()
'* Purpose: Gets direct members of a group or container object.
'* Input:   objADs          an ADs object
'*          objSchema       the schema object of objADs
'*          strClass        class name of the member objects
'*          objOutputFile   an output file object
'* Output:  The domain names are either printed on screen or saved in objOutputFile.
'*
'********************************************************************

Private Function intGetDirectMembers(objADs, objSchema, strClass, objOutputFile)

    ON ERROR RESUME NEXT

    Dim objMember, strMessage, strObjType, i

    'Initialize variables
	intGetDirectMembers = 0
	i = 0
    strMessage = ""
    strObjType = "container"    'Default this to a container object

    If IsObject(objSchema) Then        'The schema object is received
        If objSchema.Container Then        'It's a container object
            If Not (strClass = "" OR strClass = "*") Then    'If strClass is specified
                objADs.Filter = Array(CStr(strClass))
            End If
            For Each objMember in objADs
                If Err.Number then
                    Err.Clear
                Else
                    i = i + 1
                    If strClass = "" Then
                        strMessage = objMember.ADsPath
                    Else
                        strMessage = strClass & " " & i & "     " & objMember.ADsPath
                    End If
                    WriteLine strMessage, objOutputFile
                End If
            Next
        Else							'It's a group object
            strClass = ""				'Reset the class string since it can not be used.
            strObjType = "group"
            For Each objMember in objADs.Members
                If Err.Number then
                    strObjType = "leaf"
                    Err.Clear
                    Exit For
                Else
                    i = i + 1
                    If strClass = "" Then
                        strMessage = objMember.ADsPath
                    Else
                        strMessage = strClass & " " & i & "     " & objMember.ADsPath
                    End If
                    WriteLine strMessage, objOutputFile
                End If
            Next
        End If
    Else            'Could not find the schema object
        strClass = ""        'Reset the class string since it can not be used.
        'First treat it like a container
        i = 0
        For Each objMember in objADs
            If Err.Number Then
                Exit For
            End If
            i = i + 1
            strMessage = objMember.ADsPath
            WriteLine strMessage, objOutputFile
        Next
        If Err.Number then                    'It is not a container object
            Err.Clear
            i = 0
            For Each objMember in objADs.Members      'now treat it like a group object
                strObjType = "group"
                If Err.Number Then
                    strObjType = "leaf"
                    Err.Clear
                    Exit For
                End If
                i = i + 1
                strMessage = objMember.ADsPath
                WriteLine strMessage, objOutputFile
            Next
        End If
    End If

	intGetDirectMembers = i

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
'* Purpose: Prints a message on screen
'* Input:   strMessage - the string to print
'* Output:  strMessage is printed on screen
'*
'********************************************************************

Sub Print (ByRef strMessage)
    Wscript.Echo strMessage
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
'* Procedures calling sequence: CLASSIFYMEMBERS.VBS
'*
'*  intChkProgram
'*  intParseCmdLine
'*  ShowUsage
'*  GetMembers
'*      GetDirectMembers
'*          WriteLine
'*
'********************************************************************
