'********************************************************************
'*
'* File:           Share.vbs
'* Created:        March 1999
'* Version:        1.0
'*
'*  Main Function:  Lists, creates, or deletes shares from a machine.
'*
'*    1.  Share.vbs /L [/S <server>] [/U <username>]
'*                 [/W <password>]
'*                 [/O <outputfile>]
'*    2.  Share.vbs /C /N <name> /P <path> [/T <type>] [/V <description>]
'*                 [/S <server>] [/U <username>] [/W <password>]
'*                 [/O <outputfile>]
'*    3.  Share.vbs /D /N <name>
'*                 [/S <server>] [/U <username>] [/W <password>]
'*                 [/O <outputfile>]
'*
'* Copyright (C) 1999 Microsoft Corporation
'*
'********************************************************************

OPTION EXPLICIT

    'Define constants
    CONST CONST_ERROR                   = 0
    CONST CONST_WSCRIPT                 = 1
    CONST CONST_CSCRIPT                 = 2
    CONST CONST_SHOW_USAGE              = 3
    CONST CONST_PROCEED                 = 4
    CONST CONST_DELETE                  = "DELETE"
    CONST CONST_LIST                    = "LIST"
    CONST CONST_CREATE                  = "CREATE"
    
    'Declare variables
    Dim intOpMode, i
    Dim strServer, strUserName, strPassword, strOutputFile
    Dim strTaskCommand, strShareName, strSharePath
    Dim strShareComment, strShareType

    'Make sure the host is csript, if not then abort
    VerifyHostIsCscript()

    'Parse the command line
    intOpMode = intParseCmdLine(strServer       , _
                                strUserName     , _
                                strPassword     , _
                                strOutputFile   , _
                                strTaskCommand  , _
                                strShareName    , _
                                strSharePath    , _
                                strShareType    , _
                                strShareComment   )


    Select Case intOpMode

        Case CONST_SHOW_USAGE
            Call ShowUsage()

	    Case CONST_PROCEED		
		    Call Share(strServer       , _
                       strUserName     , _
                       strPassword     , _
                       strOutputFile   , _
                       strTaskCommand  , _
                       strShareName    , _
                       strSharePath    , _
                       strShareType    , _
                       strShareComment   )

        Case CONST_ERROR
            Call Wscript.Echo("Error occurred in passing parameters.")

        Case Else                    'Default -- should never happen
            Call Wscript.Echo("Error occurred in passing parameters.")

    End Select

'********************************************************************
'* End of Script
'********************************************************************

'********************************************************************
'*
'* Sub Share()
'* Purpose:	Lists, creates, or deletes shares from a machine.
'* Input:	strServer       name of the machine to be checked
'*          strTaskCommand one of list, create, and delete
'*          strShareName    name of the share to be created or deleted
'*          strSharePath    path of the share to be created
'*          strShareType    type of the share to be created
'*          strShareComment a comment for the share to be created
'*			strUserName		the current user's name
'*			strPassword		the current user's password
'*			strOutputFile	an output file name
'* Output:	Results are either printed on screen or saved in strOutputFile.
'*
'********************************************************************

Private Sub Share(strServer, strUserName, strPassword, strOutputFile, _
    strTaskCommand,strShareName,strSharePath,strShareType,strShareComment )

    ON ERROR RESUME NEXT

    Dim objFileSystem, objOutputFile, objService, strQuery


   'Open a text file for output if the file is requested
    If Not IsEmpty(strOutputFile) Then
        If (NOT blnOpenFile(strOutputFile, objOutputFile)) Then
            Call Wscript.Echo ("Could not open an output file.")
            Exit Sub
        End If
    End If

    'Establish a connection with the server.
    If blnConnect("root\cimv2" , _
                   strUserName , _
                   strPassword , _
                   strServer   , _
                   objService  ) Then
        Call Wscript.Echo("")
        Call Wscript.Echo("Please check the server name, " _
                        & "credentials and WBEM Core.")
        Exit Sub
    End If

	'Now execute the method.
	Call ExecuteMethod(objService, objOutputFile,strTaskCommand, _
         strShareName,strSharePath,strShareType,strShareComment)

    If IsObject(objOutputFile) Then
        objOutputFile.Close
        Call Wscript.Echo ("Results are saved in file " & strOutputFile & ".")
    End If

End Sub

'********************************************************************
'*
'* Sub ExecMethod()
'* Purpose:	Executes a method: creation, deletion, or listing.
'* Input:	objService	    a service object
'*		    objOutputFile	an output file object
'*          strTaskCommand one of list, create, and delete
'*          strShareName    name of the share to be created or deleted
'*          strSharePath    path of the share to be created
'*          strShareType    type of the share to be created
'*          strShareComment a comment for the share to be created
'* Output:	Results are either printed on screen or saved in objOutputFile.
'*
'********************************************************************

Private Sub ExecuteMethod(objService, objOutputFile, strTaskCommand, _
            strShareName, strSharePath,strShareType,strShareComment)

    ON ERROR RESUME NEXT

    Dim intType, intShareType, i, intStatus, strMessage
    Dim objEnumerator, objInstance
    ReDim strName(0), strDescription(0), strPath(0),strType(0), intOrder(0)

    intShareType = 0
    strMessage = ""
    strName(0) = ""
    strPath(0) = ""
    strDescription(0) = ""
    strType(0) = ""
    intOrder(0) = 0

    Select Case strTaskCommand
        Case CONST_CREATE
            Set objInstance = objService.Get("Win32_Share")
            If Err.Number Then
                Print "Error 0x" & CStr(Hex(Err.Number)) & _
                  " occurred in getting " & " a share object."
                If Err.Description <> "" Then
                     Print "Error description: " & Err.Description & "."
                End If
                Err.Clear
                Exit Sub
            End If

            If objInstance is nothing Then
                Exit Sub
            Else
                Select Case strShareType
                    Case "Disk"
                        intShareType = 0
                    Case "PrinterQ"
                        intShareType = 1
                    Case "Device"
                        intShareType = 2
                    Case "IPC"
                        intShareType = 3
                    Case "Disk$"
                        intShareType = -2147483648
                    Case "PrinterQ$"
                        intShareType = -2147483647
                    Case "Device$"
                        intShareType = -2147483646
                    Case "IPC$"
                        intShareType = -2147483645
                End Select

                intStatus = objInstance.Create(strSharePath, strShareName, _
                    intShareType, null, strShareComment, null, null)
                If intStatus = 0 Then
                    strMessage = "Succeeded in creating share " & _
                      strShareName & "."
                Else
                    strMessage = "Failed to create share " & strShareName & "."
                    strMessage = strMessage & vbCRLF & "Status = " & _
                      intStatus & "."
                End If

                WriteLine strMessage, objOutputFile
                i = i + 1
            End If
        Case CONST_DELETE
            Set objInstance = objService.Get("Win32_Share='" & strShareName _
              & "'")
            If Err.Number Then
                Print "Error 0x" & CStr(Hex(Err.Number)) & _
                  " occurred in getting share " _
                    & strShareName & "."
                If Err.Description <> "" Then
                    Print "Error description: " & Err.Description & "."
                End If
                Err.Clear
                Exit Sub
            End If

            If objInstance is nothing Then
                Exit Sub
            Else
                intStatus = objInstance.Delete()
                If intStatus = 0 Then
                    strMessage = "Succeeded in deleting share " & _
                    strShareName & "."
                Else
                    strMessage = "Failed to delete share " & strShareName & "."
                    strMessage = strMessage & vbCRLF & "Status = " & _
                    intStatus & "."
                End If
                WriteLine strMessage, objOutputFile
                i = i + 1
            End If
        Case CONST_LIST
            Set objEnumerator = objService.ExecQuery (_
                "Select Name,Description,Path,Type From Win32_Share",,0)
            If Err.Number Then
                Print "Error 0x" & CStr(Hex(Err.Number)) & _
                  " occurred during the query."
                If Err.Description <> "" Then
                    Print "Error description: " & Err.Description & "."
                End If
                Err.Clear
                Exit Sub
            End If
            Call WriteLine("There are " & objEnumerator.Count & _
              " shares.", objOutputFile)
            Call WriteLine("",objOutputFile)
            For Each objInstance in objEnumerator
                I = I + 1
                If objInstance is nothing Then
                    Exit Sub
                End If
                Call WriteLine("Name        : " & objInstance.Name, _
                     objOutputFile)
                Select Case objInstance.Type
                    Case 0
                        Call WriteLine("Type        : " & "Disk", _
                             objOutputFile)
                    Case 1
                        Call WriteLine("Type        : " & "PrinterQ", _
                             objOutputFile)
                    Case 2
                        Call WriteLine("Type        : " & "Device", _
                             objOutputFile)
                    Case 3
                        Call WriteLine("Type        : " & "IPC", _
                             objOutputFile)
                    Case -2147483648
                        Call WriteLine("Type        : " & "Disk$", _
                             objOutputFile)
                    Case -2147483647
                        Call WriteLine("Type        : " & "PrinterQ$", _
                             objOutputFile)
                    Case -2147483646
                        Call WriteLine("Type        : " & "Device$", _
                             objOutputFile)
                    Case -2147483645
                        Call WriteLine("Type        : " & "IPC$", _
                             objOutputFile)
                    Case else
                        Call WriteLine("Type        : " & "Unknown", _
                             objOutputFile)
                        strType (i) = "Unknown"
                End Select
                Call WriteLine("Description : " & _
                     objInstance.Description, objOutputFile)
                Call WriteLine("Path        : " & _
                     objInstance.Path, objOutputFile)
                Call WriteLine("", objOutputFile)
            Next

            If i > 0 Then

            Else
                strMessage = "No share is found."
                WriteLine strMessage, objOutputFile
           End If
    End Select

End Sub

'********************************************************************
'*
'* Function intParseCmdLine()
'*
'* Purpose: Parses the command line.
'* Input:   
'*
'* Output:  strServer         a remote server ("" = local server")
'*          strUserName       the current user's name
'*          strPassword       the current user's password
'*          strOutputFile     an output file name
'*
'********************************************************************
Private Function intParseCmdLine(strServer       , _
                                 strUserName     , _
                                 strPassword     , _
                                 strOutputFile   , _
                                 strTaskCommand  , _
                                 strShareName    , _
                                 strSharePath    , _
                                 strShareType    , _
                                 strShareComment   )

    ON ERROR RESUME NEXT

    Dim strFlag
    Dim intState, intArgIter
    Dim objFileSystem

    If Wscript.Arguments.Count > 0 Then
        strFlag = Wscript.arguments.Item(0)
    End If

    If IsEmpty(strFlag) Then                'No arguments have been received
        intParseCmdLine = CONST_PROCEED
        strTaskCommand = CONST_LIST
        Exit Function
    End If

    'Check if the user is asking for help or is just confused
    If (strFlag="help") OR (strFlag="/h") OR (strFlag="\h") OR (strFlag="-h") _
        OR (strFlag = "\?") OR (strFlag = "/?") OR (strFlag = "?") _ 
        OR (strFlag="h") Then
        intParseCmdLine = CONST_SHOW_USAGE
        Exit Function
    End If

    'Retrieve the command line and set appropriate variables
     intArgIter = 0
    Do While intArgIter <= Wscript.arguments.Count - 1
        Select Case Left(LCase(Wscript.arguments.Item(intArgIter)),2)
  
            Case "/s"
                If Not blnGetArg("Server", strServer, intArgIter) Then
                    intParseCmdLine = CONST_ERROR
                    Exit Function
                End If
                intArgIter = intArgIter + 1

            Case "/o"
                If Not blnGetArg("Output File", strOutputFile, intArgIter) Then
                    intParseCmdLine = CONST_ERROR
                    Exit Function
                End If
                intArgIter = intArgIter + 1

            Case "/u"
                If Not blnGetArg("User Name", strUserName, intArgIter) Then
                    intParseCmdLine = CONST_ERROR
                    Exit Function
                End If
                intArgIter = intArgIter + 1

            Case "/w"
                If Not blnGetArg("User Password", strPassword, intArgIter) Then
                    intParseCmdLine = CONST_ERROR
                    Exit Function
                End If
                intArgIter = intArgIter + 1

            Case "/l"
                intParseCmdLine = CONST_PROCEED
                strTaskCommand = CONST_LIST
                intArgITer = intArgITer + 1
               
            Case "/c"
                intParseCmdLine = CONST_PROCEED
                strTaskCommand = CONST_CREATE
                intArgITer = intArgITer + 1

            Case "/d"
                intParseCmdLine = CONST_PROCEED
                strTaskCommand = CONST_DELETE
                intArgITer = intArgITer + 1

            Case "/n"
                If Not blnGetArg("Share Name", strShareName, intArgIter) Then
                    intParseCmdLine = CONST_ERROR
                    Exit Function
                End If
                intArgITer = intArgITer + 1

            Case "/p"
                If Not blnGetArg("Share Path", strSharePath, intArgIter) Then
                    intParseCmdLine = CONST_ERROR
                    Exit Function
                End If
                intArgITer = intArgITer + 1

            Case "/t"
                If Not blnGetArg("Share Type", strShareType, intArgIter) Then
                    intParseCmdLine = CONST_ERROR
                    Exit Function
                End If
                intArgITer = intArgITer + 1

            Case "/v"
                If Not blnGetArg("Share Description", strShareComment, _
                    intArgIter) Then
                    intParseCmdLine = CONST_ERROR
                    Exit Function
                End If
                intArgITer = intArgITer + 1

            Case Else 'We shouldn't get here
                Call Wscript.Echo("Invalid or misplaced parameter: " _
                   & Wscript.arguments.Item(intArgIter) & vbCRLF _
                   & "Please check the input and try again," & vbCRLF _
                   & "or invoke with '/?' for help with the syntax.")
                Wscript.Quit

        End Select

    Loop '** intArgIter <= Wscript.arguments.Count - 1

    If IsEmpty(intParseCmdLine) Then
        intParseCmdLine = CONST_PROCEED
        strTaskCommand = CONST_LIST
    End If
    Select Case strTaskCommand
        Case CONST_CREATE
            If IsEmpty(strShareName) then
                intParseCmdLine = CONST_ERROR
            End IF
            If IsEmpty(strSharePath) then
                intParseCmdLine = CONST_ERROR
            End If
        Case CONST_DELETE
            If IsEmpty(strShareName) then
                intParseCmdLine = CONST_ERROR
            End IF
    End Select

End Function

'********************************************************************
'*
'* Sub ShowUsage()
'*
'* Purpose: Shows the correct usage to the user.
'*
'* Input:   None
'*
'* Output:  Help messages are displayed on screen.
'*
'********************************************************************
Private Sub ShowUsage()

    Wscript.Echo ""
    Wscript.Echo "Lists, creates, or deletes shares from a machine."
    Wscript.Echo ""
    Wscript.Echo "SYNTAX:"
    Wscript.Echo "1.  Share.vbs /L [/S <server>] [/U <username>]" _
                &" [/W <password>]"
    Wscript.Echo "                 [/O <outputfile>]"
    Wscript.Echo "2.  Share.vbs /C /N <name> /P <path> [/T <type>]" _
               & " [/V <description>]"
    Wscript.Echo "             [/S <server>] [/U <username>] [/W <password>]"
    Wscript.Echo "             [/O <outputfile>]"
    Wscript.Echo "3.  Share.vbs /D /N <name>"
    Wscript.Echo "             [/S <server>] [/U <username>] [/W <password>]"
    Wscript.Echo "             [/O <outputfile>]"
    Wscript.Echo ""
    Wscript.Echo "PARAMETER SPECIFIERS:"
    Wscript.Echo "   /L            Lists all shares on a machine."
    Wscript.Echo "   /C            Creates a share on a machine."
    Wscript.Echo "   /D            Deletes a share from a machine."
    Wscript.echo "   name          Name of the share to be created or deleted."
    Wscript.echo "   path          Path of the share to be created."
    Wscript.Echo "   description   A description for the share."
    Wscript.echo "   type          Type of the share to be created. Must be one"
    Wscript.echo "                 one of Disk, Printer, IPC, Special."
    Wscript.Echo "   server        A machine name."
    Wscript.Echo "   username      The current user's name."
    Wscript.Echo "   password      Password of the current user."
    Wscript.Echo "   outputfile    The output file name."
    Wscript.Echo ""
    Wscript.Echo "EXAMPLE:"
    Wscript.Echo "1. cscript Share.vbs /l /s MyMachine2"
    Wscript.Echo "   List the shares on the machine MyMachine2."
    Wscript.Echo "2. cscript Share.vbs /c /n scratch /p c:\scratch " _
               & "/t Disk /v ""Scratch Directory"""
    Wscript.Echo "   Creates a file share called ""scratch"" on the" _
               & " local machine."
    Wscript.Echo "3. cscript Share.vbs /d /n scratch /s MyMachine2."
    Wscript.Echo "   Deletes the share named ""scratch"" on the machine" _
               & " MyMachine2."

End Sub


'********************************************************************
'*
'* Sub SortArray()
'* Purpose: Sorts an array and arrange another array accordingly.
'* Input:   strArray    the array to be sorted
'*          blnOrder    True for ascending and False for descending
'*          strArray2   an array that has exactly the same number of
'*                      elements as strArray
'*                      and will be reordered together with strArray
'*          blnCase     indicates whether the order is case sensitive
'* Output:  The sorted arrays are returned in the original arrays.
'* Note:    Repeating elements are not deleted.
'*
'********************************************************************

Private Sub SortArray(strArray, blnOrder, strArray2, blnCase)

    ON ERROR RESUME NEXT

    Dim i, j, intUbound

    If IsArray(strArray) Then
        intUbound = UBound(strArray)
    Else
        Print "Argument is not an array!"
        Exit Sub
    End If

    blnOrder = CBool(blnOrder)
    blnCase = CBool(blnCase)
    If Err.Number Then
        Print "Argument is not a boolean!"
        Exit Sub
    End If

    i = 0
    Do Until i > intUbound-1
        j = i + 1
        Do Until j > intUbound
            If blnCase Then     'Case sensitive
                If (strArray(i) > strArray(j)) and blnOrder Then
                    Swap strArray(i), strArray(j)   'swaps element i and j
                    Swap strArray2(i), strArray2(j)
                ElseIf (strArray(i) < strArray(j)) and Not blnOrder Then
                    Swap strArray(i), strArray(j)   'swaps element i and j
                    Swap strArray2(i), strArray2(j)
                ElseIf strArray(i) = strArray(j) Then
                    'Move element j to next to i
                    If j > i + 1 Then
                        Swap strArray(i+1), strArray(j)
                        Swap strArray2(i+1), strArray2(j)
                    End If
                End If
            Else
                If (LCase(strArray(i)) > LCase(strArray(j))) and blnOrder Then
                    Swap strArray(i), strArray(j)   'swaps element i and j
                    Swap strArray2(i), strArray2(j)
                ElseIf (LCase(strArray(i)) < LCase(strArray(j))) _
                    and Not blnOrder Then
                    Swap strArray(i), strArray(j)   'swaps element i and j
                    Swap strArray2(i), strArray2(j)
                ElseIf LCase(strArray(i)) = LCase(strArray(j)) Then
                    'Move element j to next to i
                    If j > i + 1 Then
                        Swap strArray(i+1), strArray(j)
                        Swap strArray2(i+1), strArray2(j)
                    End If
                End If
            End If
            j = j + 1
        Loop
        i = i + 1
    Loop

End Sub

'********************************************************************
'*
'* Sub Swap()
'* Purpose: Exchanges values of two strings.
'* Input:   strA    a string
'*          strB    another string
'* Output:  Values of strA and strB are exchanged.
'*
'********************************************************************

Private Sub Swap(ByRef strA, ByRef strB)

    Dim strTemp

    strTemp = strA
    strA = strB
    strB = strTemp

End Sub

'********************************************************************
'*
'* Sub ReArrangeArray()
'* Purpose: Rearranges one array according to order specified in another array.
'* Input:   strArray    the array to be rearranged
'*          intOrder    an integer array that specifies the order
'* Output:  strArray is returned as rearranged
'*
'********************************************************************

Private Sub ReArrangeArray(strArray, intOrder)

    ON ERROR RESUME NEXT

    Dim intUBound, i, strTempArray()

    If Not (IsArray(strArray) and IsArray(intOrder)) Then
        Print "At least one of the arguments is not an array"
        Exit Sub
    End If

    intUBound = UBound(strArray)

    If intUBound <> UBound(intOrder) Then
        Print "The upper bound of these two arrays do not match!"
        Exit Sub
    End If

    ReDim strTempArray(intUBound)

    For i = 0 To intUBound
        strTempArray(i) = strArray(intOrder(i))
        If Err.Number Then
            Print "Error 0x" & CStr(Hex(Err.Number)) & _
                " occurred in rearranging an array."
            If Err.Description <> "" Then
                Print "Error description: " & Err.Description & "."
            End If
            Err.Clear
            Exit Sub
        End If
    Next

    For i = 0 To intUBound
        strArray(i) = strTempArray(i)
    Next

End Sub

'********************************************************************
'* General Routines
'********************************************************************

'********************************************************************
'*
'* Function strPackString()
'*
'* Purpose: Attaches spaces to a string to increase the length to intWidth.
'*
'* Input:   strString   a string
'*          intWidth    the intended length of the string
'*          blnAfter    Should spaces be added after the string?
'*          blnTruncate specifies whether to truncate the string or not if
'*                      the string length is longer than intWidth
'*
'* Output:  strPackString is returned as the packed string.
'*
'********************************************************************
Private Function strPackString( ByVal strString, _
                                ByVal intWidth,  _
                                ByVal blnAfter,  _
                                ByVal blnTruncate)

    ON ERROR RESUME NEXT

    intWidth      = CInt(intWidth)
    blnAfter      = CBool(blnAfter)
    blnTruncate   = CBool(blnTruncate)

    If Err.Number Then
        Call Wscript.Echo ("Argument type is incorrect!")
        Err.Clear
        Wscript.Quit
    End If

    If IsNull(strString) Then
        strPackString = "null" & Space(intWidth-4)
        Exit Function
    End If

    strString = CStr(strString)
    If Err.Number Then
        Call Wscript.Echo ("Argument type is incorrect!")
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
'*  Function blnGetArg()
'*
'*  Purpose: Helper to intParseCmdLine()
'* 
'*  Usage:
'*
'*     Case "/s" 
'*       blnGetArg ("server name", strServer, intArgIter)
'*
'********************************************************************
Private Function blnGetArg ( ByVal StrVarName,   _
                             ByRef strVar,       _
                             ByRef intArgIter) 

    blnGetArg = False 'failure, changed to True upon successful completion

    If Len(Wscript.Arguments(intArgIter)) > 2 then
        If Mid(Wscript.Arguments(intArgIter),3,1) = ":" then
            If Len(Wscript.Arguments(intArgIter)) > 3 then
                strVar = Right(Wscript.Arguments(intArgIter), _
                         Len(Wscript.Arguments(intArgIter)) - 3)
                blnGetArg = True
                Exit Function
            Else
                intArgIter = intArgIter + 1
                If intArgIter > (Wscript.Arguments.Count - 1) Then
                    Call Wscript.Echo( "Invalid " & StrVarName & ".")
                    Call Wscript.Echo( "Please check the input and try again.")
                    Exit Function
                End If

                strVar = Wscript.Arguments.Item(intArgIter)
                If Err.Number Then
                    Call Wscript.Echo( "Invalid " & StrVarName & ".")
                    Call Wscript.Echo( "Please check the input and try again.")
                    Exit Function
                End If

                If InStr(strVar, "/") Then
                    Call Wscript.Echo( "Invalid " & StrVarName)
                    Call Wscript.Echo( "Please check the input and try again.")
                    Exit Function
                End If

                blnGetArg = True 'success
            End If
        Else
            strVar = Right(Wscript.Arguments(intArgIter), _
                     Len(Wscript.Arguments(intArgIter)) - 2)
            blnGetArg = True 'success
            Exit Function
        End If
    Else
        intArgIter = intArgIter + 1
        If intArgIter > (Wscript.Arguments.Count - 1) Then
            Call Wscript.Echo( "Invalid " & StrVarName & ".")
            Call Wscript.Echo( "Please check the input and try again.")
            Exit Function
        End If

        strVar = Wscript.Arguments.Item(intArgIter)
        If Err.Number Then
            Call Wscript.Echo( "Invalid " & StrVarName & ".")
            Call Wscript.Echo( "Please check the input and try again.")
            Exit Function
        End If

        If InStr(strVar, "/") Then
            Call Wscript.Echo( "Invalid " & StrVarName)
            Call Wscript.Echo( "Please check the input and try again.")
            Exit Function
        End If
        blnGetArg = True 'success
    End If
End Function

'********************************************************************
'*
'* Function blnConnect()
'*
'* Purpose: Connects to machine strServer.
'*
'* Input:   strServer       a machine name
'*          strNameSpace    a namespace
'*          strUserName     name of the current user
'*          strPassword     password of the current user
'*
'* Output:  objService is returned  as a service object.
'*          strServer is set to local host if left unspecified
'*
'********************************************************************
Private Function blnConnect(ByVal strNameSpace, _
                            ByVal strUserName,  _
                            ByVal strPassword,  _
                            ByRef strServer,    _
                            ByRef objService)

    ON ERROR RESUME NEXT

    Dim objLocator, objWshNet

    blnConnect = False     'There is no error.

    'Create Locator object to connect to remote CIM object manager
    Set objLocator = CreateObject("WbemScripting.SWbemLocator")
    If Err.Number then
        Call Wscript.Echo( "Error 0x" & CStr(Hex(Err.Number)) & _
                           " occurred in creating a locator object." )
        If Err.Description <> "" Then
            Call Wscript.Echo( "Error description: " & Err.Description & "." )
        End If
        Err.Clear
        blnConnect = True     'An error occurred
        Exit Function
    End If

    'Connect to the namespace which is either local or remote
    Set objService = objLocator.ConnectServer (strServer, strNameSpace, _
       strUserName, strPassword)
    ObjService.Security_.impersonationlevel = 3
    If Err.Number then
        Call Wscript.Echo( "Error 0x" & CStr(Hex(Err.Number)) & _
                           " occurred in connecting to server " _
           & strServer & ".")
        If Err.Description <> "" Then
            Call Wscript.Echo( "Error description: " & Err.Description & "." )
        End If
        Err.Clear
        blnConnect = True     'An error occurred
    End If

    'Get the current server's name if left unspecified
    If IsEmpty(strServer) Then
        Set objWshNet = CreateObject("Wscript.Network")
    strServer     = objWshNet.ComputerName
    End If

End Function

'********************************************************************
'*
'* Sub      VerifyHostIsCscript()
'*
'* Purpose: Determines which program is used to run this script.
'*
'* Input:   None
'*
'* Output:  If host is not cscript, then an error message is printed 
'*          and the script is aborted.
'*
'********************************************************************
Sub VerifyHostIsCscript()

    ON ERROR RESUME NEXT

    Dim strFullName, strCommand, i, j, intStatus

    strFullName = WScript.FullName

    If Err.Number then
        Call Wscript.Echo( "Error 0x" & CStr(Hex(Err.Number)) & " occurred." )
        If Err.Description <> "" Then
            Call Wscript.Echo( "Error description: " & Err.Description & "." )
        End If
        intStatus =  CONST_ERROR
    End If

    i = InStr(1, strFullName, ".exe", 1)
    If i = 0 Then
        intStatus =  CONST_ERROR
    Else
        j = InStrRev(strFullName, "\", i, 1)
        If j = 0 Then
            intStatus =  CONST_ERROR
        Else
            strCommand = Mid(strFullName, j+1, i-j-1)
            Select Case LCase(strCommand)
                Case "cscript"
                    intStatus = CONST_CSCRIPT
                Case "wscript"
                    intStatus = CONST_WSCRIPT
                Case Else       'should never happen
                    Call Wscript.Echo( "An unexpected program was used to " _
                                       & "run this script." )
                    Call Wscript.Echo( "Only CScript.Exe or WScript.Exe can " _
                                       & "be used to run this script." )
                    intStatus = CONST_ERROR
                End Select
        End If
    End If

    If intStatus <> CONST_CSCRIPT Then
        Call WScript.Echo( "Please run this script using CScript." & vbCRLF & _
             "This can be achieved by" & vbCRLF & _
             "1. Using ""CScript Share.vbs arguments"" for Windows 95/98 or" _
             & vbCRLF & "2. Changing the default Windows Scripting Host " _
             & "setting to CScript" & vbCRLF & "    using ""CScript " _
             & "//H:CScript //S"" and running the script using" & vbCRLF & _
             "    ""Share.vbs arguments"" for Windows NT/2000." )
        WScript.Quit
    End If

End Sub

'********************************************************************
'*
'* Sub WriteLine()
'* Purpose: Writes a text line either to a file or on screen.
'* Input:   strMessage  the string to print
'*          objFile     an output file object
'* Output:  strMessage is either displayed on screen or written to a file.
'*
'********************************************************************
Sub WriteLine(ByVal strMessage, ByVal objFile)

    On Error Resume Next
    If IsObject(objFile) then        'objFile should be a file object
        objFile.WriteLine strMessage
    Else
        Call Wscript.Echo( strMessage )
    End If

End Sub

'********************************************************************
'* 
'* Function blnErrorOccurred()
'*
'* Purpose: Reports error with a string saying what the error occurred in.
'*
'* Input:   strIn		string saying what the error occurred in.
'*
'* Output:  displayed on screen 
'* 
'********************************************************************
Private Function blnErrorOccurred (ByVal strIn)

    If Err.Number Then
        Call Wscript.Echo( "Error 0x" & CStr(Hex(Err.Number)) & ": " & strIn)
        If Err.Description <> "" Then
            Call Wscript.Echo( "Error description: " & Err.Description)
        End If
        Err.Clear
        blnErrorOccurred = True
    Else
        blnErrorOccurred = False
    End If

End Function

'********************************************************************
'* 
'* Function blnOpenFile
'*
'* Purpose: Opens a file.
'*
'* Input:   strFileName		A string with the name of the file.
'*
'* Output:  Sets objOpenFile to a FileSystemObject and setis it to 
'*            Nothing upon Failure.
'* 
'********************************************************************
Private Function blnOpenFile(ByVal strFileName, ByRef objOpenFile)

    ON ERROR RESUME NEXT

    Dim objFileSystem

    Set objFileSystem = Nothing

    If IsEmpty(strFileName) OR strFileName = "" Then
        blnOpenFile = False
        Set objOpenFile = Nothing
        Exit Function
    End If

    'Create a file object
    Set objFileSystem = CreateObject("Scripting.FileSystemObject")
    If blnErrorOccurred("Could not create filesystem object.") Then
        blnOpenFile = False
        Set objOpenFile = Nothing
        Exit Function
    End If

    'Open the file for output
    Set objOpenFile = objFileSystem.OpenTextFile(strFileName, 8, True)
    If blnErrorOccurred("Could not open") Then
        blnOpenFile = False
        Set objOpenFile = Nothing
        Exit Function
    End If
    blnOpenFile = True

End Function

'********************************************************************
'*                                                                  *
'*                           End of File                            *
'*                                                                  *
'********************************************************************


