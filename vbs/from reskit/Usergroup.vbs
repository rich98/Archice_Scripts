
'********************************************************************
'* File:          USERGROUP.VBS
'* Created:       May 1998
'* Version:       1.0
'*
'* Main Function: Adds or deletes one or multiple users to or from a group.
'* Usage: USERGROUP.VBS grouppath </A | /D | /L> </A:userpath | /I:inputfile>
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
CONST CONST_ADD						= "add"
CONST CONST_DELETE					= "delete"
CONST CONST_LIST					= "list"

'Declare variables
Dim strInputFile, intOpMode, blnQuiet, strAction, i
Dim strGroupPath, strUserName, strPassword
ReDim strUserPaths(0), strArgumentArray(0)

'Initialize variables
strGroupPath = ""
strUserName = ""
strPassword = ""
strInputFile = ""
blnQuiet = False
strAction = CONST_ADD		'Default to add users to group.
strArgumentArray(0) = ""
strUserPaths(0) = ""

'Get the command line arguments
For i = 0 to Wscript.arguments.count - 1
    ReDim Preserve strArgumentArray(i)
    strArgumentArray(i) = Wscript.arguments.Item(i)
Next

'Check whether the script is run using CScript
Select Case intChkProgram()
    Case CONST_CSCRIPT
        'Do Nothing
    Case CONST_WSCRIPT
        WScript.Echo "Please run this script using CScript." & vbCRLF & _
            "This can be achieved by" & vbCRLF & _
            "1. Using ""CScript UserGroup.vbs arguments"" for Windows 95/98 or" & vbCRLF & _
            "2. Changing the default Windows Scripting Host setting to CScript" & vbCRLF & _
            "    using ""CScript //H:CScript //S"" and running the script using" & vbCRLF & _
            "    ""UserGroup.vbs arguments"" for Windows NT."
        WScript.Quit
    Case Else
        WScript.Quit
End Select

'Parse the command line
intOpMode = intParseCmdLine(strArgumentArray, strGroupPath, strUserPaths, _
            strInputFile, strAction, strUserName, strPassword, blnQuiet)
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
        Call UserGroup(strGroupPath, strUserPaths, strInputFile, strAction, _
             strUserName, strPassword)
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
'* Output:  strGroupPath        ADsPath of a group object
'*          strUserPaths        ADsPath of a user object
'*          strInputFile        an input file name
'*          strAction           the action to take
'*          strUserName         name of the current user
'*          strPassword         password of the current user
'*          blnQuiet            specifies whether to suppress messages
'*          intParseCmdLine     is set to CONST_SHOW_USAGE if there is an error
'*                              in input and CONST_PROCEED otherwise.
'*
'********************************************************************

Private Function intParseCmdLine(strArgumentArray, strGroupPath, strUserPaths, _
    strInputFile, strAction, strUserName, strPassword, blnQuiet)

    ON ERROR RESUME NEXT

    Dim strFlag, i, intCount

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

    strGroupPath = strFlag          'The first parameter must be the group's ADsPath.
	intCount = -1
    For i = 1 to UBound(strArgumentArray)
        strFlag = LCase(Left(strArgumentArray(i), InStr(1, strArgumentArray(i), ":")-1))
        If Err.Number Then            'An error occurs if there is no : in the string
            Err.Clear
            Select Case LCase(strArgumentArray(i))
                Case "/q"
                    blnQuiet = True
                Case "/a"
                    strAction = CONST_ADD
                Case "/d"
                    strAction = CONST_DELETE
                Case "/l"
                    strAction = CONST_LIST
                Case Else
                    Print strArgumentArray(i) & " is not a valid input."
                    Print "Please check the input and try again."
                    intParseCmdLine = CONST_ERROR
                    Exit Function
            End Select
        Else
            Select Case strFlag
                Case "/a"
					intCount = intCount + 1
					ReDim Preserve strUserPaths(intCount)
                    strUserPaths(intCount) = Right(strArgumentArray(i), _
						Len(strArgumentArray(i))-3)
                Case "/i"
                    strInputFile = Right(strArgumentArray(i), Len(strArgumentArray(i))-3)
                Case "/u"
                    strUserName = Right(strArgumentArray(i), Len(strArgumentArray(i))-3)
                Case "/w"
                    strPassword = Right(strArgumentArray(i), Len(strArgumentArray(i))-3)
                Case Else
                    Print "Invalid flag " & """" & strFlag & ":""" & "."
                    Print "Please check the input and try again."
                    intParseCmdLine = CONST_ERROR
                    Exit Function
                End Select
        End If
    Next

    intParseCmdLine = CONST_PROCEED

	If strUserPaths(0) = "" And strInputFile = "" And strAction <> CONST_LIST Then
		Print "The user ADsPath is missing."
        Print "Please check the input and try again."
        intParseCmdLine = CONST_ERROR
	End If

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
    Wscript.echo "Adds or deletes one or multiple users to or from a group." & vbCRLF
    Wscript.echo "USERGROUP.VBS grouppath </A | /D | /L> </A:userpath | /I:inputfile>"
    Wscript.echo "[/U:username] [/W:password] [/Q]"
    Wscript.echo "   /A, /I, /U, /W"
    Wscript.echo "                 Parameter specifiers."
    Wscript.Echo "   grouppath     The ADsPath of a group object."
    Wscript.Echo "   /A, /D, /L    Adds users to a group, deletes users from a groups,"
    Wscript.Echo "				   or list members of a group."
    Wscript.Echo "   userpath      The ADsPath of a user object."
    Wscript.Echo "   inputfile     Name of the input file."
    Wscript.Echo "   username      Username of the current user."
    Wscript.Echo "   password      Password of the current user."
    Wscript.Echo "   /Q            Suppresses all output messages." & vbCRLF
    Wscript.Echo "EXAMPLES:"
    Wscript.echo "1. USERGROUP.VBS WinNT:\\FooFoo\administrators /A"
    Wscript.echo "   \a:WinNT:\\FooFoo1\user1 \a:WinNT:\\FooFoo2\user2"
    Wscript.echo "   adds user1 of domain FooFoo1 and user2 of domain FooFoo2 to"
    Wscript.echo "   the administrators group of domain FooFoo."
    Wscript.echo "2. USERGROUP.VBS WinNT:\\FooFoo\administrators /D" 
    Wscript.echo "   \a:WinNT:\\FooFoo2\user2"
    Wscript.echo "   deletes user2 of domain FooFoo2 from the administrators group of"
    Wscript.echo "   domain FooFoo."

End Sub

'********************************************************************
'*
'* Sub UserGroup()
'* Purpose: Adds or deletes one or multiple users to or from a group.
'* Input:   strGroupPath        ADsPath of a group object
'*          strUserPaths        ADsPath of a user object
'*          strInputFile        an input file name
'*          strAction           deletes user(s) from a group
'*          strUserName         name of the current user
'*          strPassword         password of the current user
'* Output:  Results are either printed on screen or saved in strInputFile.
'*
'********************************************************************

Private Sub UserGroup(strGroupPath, strUserPaths, strInputFile, strAction, _
    strUserName, strPassword)

    ON ERROR RESUME NEXT

    Dim objGroup, objUser, objProvider, strProvider, strUserPath
	Dim objFileSystem, objInputFile, strMessage, i, intFlag

	Print "Getting object " & strGroupPath & "..."
	If strUserName = ""	then		'The current user is assumed
		set objGroup = GetObject(strGroupPath)
	Else						'Credentials are passed
		strProvider = Left(strGroupPath, InStr(1, strGroupPath, ":"))
		set objProvider = GetObject(strProvider)
        'Use user authentication
		set objGroup = objProvider.OpenDsObject(strGroupPath,strUserName,strPassword,1)		
	End If
    If Err.Number then
		If CStr(Hex(Err.Number)) = "80070035" Then
			Print "Object " & strGroupPath & " is not found."
		Else
			Print "Error 0x" & CStr(Hex(Err.Number)) & " occurred in getting object " _
				& strGroupPath & "."
			If Err.Description <> "" Then
				Print "Error description: " & Err.Description & "."
			End If
		End If
		Err.Clear
        Exit Sub
    End If

	If strAction = CONST_LIST Then
		Print "Members of " & objGroup.ADsPath & ":"
		i = 0
		For Each objUser In objGroup.Members
			Print "    " & objUser.ADsPath
			i = i + 1
		Next
		If i = 0 Then
			Print objGroup.ADsPath & " has no members!"

		End If
		Exit Sub
	End If

	intFlag = 0
    If strUserPaths(0) <> "" Then
		For i = 0 To UBound(strUserPaths)
			Select Case strAction 
				Case CONST_ADD
					If objGroup.IsMember(strUserPaths(i)) Then
						Print strUserPaths(i) & " is already a member of " _
							& objGroup.ADsPath & "."
						intFlag = 0
					Else
						objGroup.Add(strUserPaths(i))
						intFlag = 1
					End If
				Case CONST_DELETE
					If objGroup.IsMember(strUserPaths(i)) Then
						objGroup.Remove(strUserPaths(i))
						intFlag = 1
					Else
						Print strUserPaths(i) & " is not a member of " _
							& objGroup.ADsPath & "."
						intFlag = 0
					End If
				Case Else
					Print "Action " & strAction & " is unknown!"
					Exit Sub
			End Select
			If Err.Number Then
				If strAction = CONST_DELETE Then
					strMessage = " occurred in deleting user " & strUserPaths(i) _
						& " from group " & strGroupPath & "."
				ElseIf strAction = CONST_ADD Then
					strMessage = " occurred in adding user " & strUserPaths(i) _
						& " to group " & strGroupPath & "."
				End If
				Print "Error 0x" & CStr(Hex(Err.Number)) & strMessage
				If Err.Description <> "" Then
					Print "Error description: " & Err.Description & "."
				End If
				Print "Clear error and continue."
				Err.Clear
			ElseIf intFlag = 1 Then
				If strAction = CONST_DELETE Then
					strMessage = "Succeeded in deleting user " & strUserPaths(i) _
					& " from group " & strGroupPath & "."
				ElseIf strAction = CONST_ADD Then
					strMessage = "Succeeded in adding user " & strUserPaths(i) _
					& " to group " & strGroupPath & "."
				End If
				Print strMessage
			End If
		Next
    End If

    If strInputFile <> "" Then
        'Create a filesystem object.
        Set objFileSystem = CreateObject("Scripting.FileSystemObject")
        If Err.Number then
            Print "Error 0x" & CStr(Hex(Err.Number)) & _
                " occurred in opening a filesystem object."
            If Err.Description <> "" Then
                Print "Error description: " & Err.Description & "."
            End If
            Exit Sub
        End If
        'Open the file For output
        Set objInputFile = objFileSystem.OpenTextFile(strInputFile)
        If Err.Number then
            Print "Error 0x" & CStr(Hex(Err.Number)) & " occurred in opening file " _
                & strInputFile
            If Err.Description <> "" Then
                Print "Error description: " & Err.Description & "."
            End If
            Exit Sub
        End If

        While Not objInputFile.AtEndOfStream
            'Read a line and get rid of leading and trailing spaces.
            strUserPath = Trim(objInputFile.ReadLine)
            If strUserPath <> "" Then
				Select Case strAction 
					Case CONST_ADD
						If objGroup.IsMember(strUserPath) Then
							Print strUserPath & " is already a member of " _
								& objGroup.ADsPath & "."
							intFlag = 0
						Else
							objGroup.Add(strUserPath)
							intFlag = 1
						End If
					Case CONST_DELETE
						If objGroup.IsMember(strUserPath) Then
							objGroup.Remove(strUserPath)
							intFlag = 1
						Else
							Print strUserPath & " is not a member of " _
								& objGroup.ADsPath & "."
							intFlag = 0
						End If
					Case Else
						Print "Action " & strAction & " is unknown!"
						Exit Sub
				End Select
			    If Err.Number Then
				    If strAction = CONST_DELETE Then
					    strMessage = " occurred in deleting user " & strUserPath _
						    & " from group " & strGroupPath & "."
				    ElseIf strAction = CONST_ADD Then
					    strMessage = " occurred in adding user " & strUserPath _
						    & " to group " & strGroupPath & "."
				    End If
				    Print "Error 0x" & CStr(Hex(Err.Number)) & strMessage
				    If Err.Description <> "" Then
					    Print "Error description: " & Err.Description & "."
				    End If
				    Print "Clear error and continue."
				    Err.Clear
			    ElseIf intFlag = 1 Then
				    If strAction = CONST_DELETE Then
					    strMessage = "Succeeded in deleting user " & strUserPath _
					    & " from group " & strGroupPath & "."
				    ElseIf strAction = CONST_ADD Then
					    strMessage = "Succeeded in adding user " & strUserPath _
					    & " to group " & strGroupPath & "."
				    End If
				    Print strMessage
			    End If
            End If
        Wend
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
'*                                                                  *
'*                           End of File                            *
'*                                                                  *
'********************************************************************

'********************************************************************
'*
'* Procedures calling sequence: USERGROUP.VBS
'*
'*  intChkProgram
'*  intParseCmdLine
'*  ShowUsage
'*  UserGroup
'*
'********************************************************************

