'********************************************************************
'*
'* File:           FileMan.vbs
'* Created:        March 1999
'* Version:        1.0
'*
'*  Main Function:  Perform various simple operations on a file.
'*
'*   1.  Fileman.vbs /D <targetfile> | /T <targetfile>
'*                  [/S <server>][/U <username>][/W <password>]
'*                  [/O <outputfile>]
'*
'*   2.  Fileman.vbs /R <targetfile> | /C <targetfile>
'*                   /N <newfile>
'*                  [/S <server>][/U <username>][/W <password>]
'*                  [/O <outputfile>]
'*
'*    Copyright (C) 1999 Microsoft Corporation
'*
'********************************************************************


OPTION EXPLICIT

    'Define constants

    CONST CONST_ERROR                   = 0
    CONST CONST_WSCRIPT                 = 1
    CONST CONST_CSCRIPT                 = 2
    CONST CONST_SHOW_USAGE              = 3
    CONST CONST_PROCEED                 = 4
    CONST CONST_DELETE                  = 5
    CONST CONST_RENAME                  = 6
    CONST CONST_COPY                    = 7
    CONST CONST_TAKEOWNERSHIP           = 8


    'Declare variables
    Dim intOpMode,   i
    Dim strServer, strUserName, strPassword, strOutputFile
    Dim strTaskCommand, strFileName, strNewFileName
    Dim blnForce

    'Make sure the host is csript, if not then abort
    VerifyHostIsCscript()

    'Parse the command line
    intOpMode = intParseCmdLine(strServer      ,  _
                                strUserName    ,  _
                                strPassword    ,  _
                                strOutputFile  ,  _
                                strTaskCommand ,  _
                                strFileName    ,  _
                                strNewFileName ,  _
                                blnForce          )

    Select Case intOpMode

        Case CONST_SHOW_USAGE
            Call ShowUsage()

        Case CONST_PROCEED
            Call FileMan(strServer      ,  _
                         strUserName    ,  _
                         strPassword    ,  _
                         strOutputFile  ,  _
                         strTaskCommand ,  _ 
                         strFileName    ,  _
                         strNewFileName ,  _
                         blnForce          )

        Case CONST_ERROR
            Call Wscript.Echo("Invalid or missing parameters " _
               & "Please check the input and try again," & vbCRLF _
               & "or invoke with '/?' for help with the syntax.")
            Wscript.Quit

        Case Else                    'Default -- should never happen
            Call Wscript.Echo("Error occurred in passing parameters.")

    End Select

'********************************************************************
'* End of Script
'********************************************************************

'********************************************************************
'*
'* Sub FileMan
'* Purpose: Perform various simple operations on a file.
'* Input:   
'*          strServer          a machine name
'*          strOutputFile      an output file name
'*          strUserName        the current user's name
'*          strPassword        the current user's password
'*          strTaskCommand     the file operation to perform 
'*          strFileName        the target file
'*          strNewFileName     the new file
'*          blnForce           overwrite read only files.
'*
'********************************************************************
Private Sub FileMan(strServer      ,  _
                    strUserName    ,  _
                    strPassword    ,  _
                    strOutputFile  ,  _
                    strTaskCommand ,  _ 
                    strFileName    ,  _
                    strNewFileName ,  _
                    blnForce          )


    ON ERROR RESUME NEXT

    Dim objFileSystem, objOutputFile, objService
    Dim strQuery


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
    Select Case strTaskCommand

        Case CONST_DELETE

            Call DeleteFile(strServer     ,  _
                            strUserName   ,  _
                            strPassword   ,  _
                            objService    ,  _
                            objOutputFile ,  _
                            strFileName   ,  _
                            blnForce         )

        Case CONST_RENAME

            Call RenameFile(strServer      ,  _
                            strUserName    ,  _
                            strPassword    ,  _
                            objService     ,  _
                            objOutputFile  ,  _
                            strFileName    ,  _
                            strNewFileName    )

        Case CONST_COPY

            Call CopyFile(strServer      ,  _
                          strUserName    ,  _
                          strPassword    ,  _
                          objService     ,  _
                          objOutputFile  ,  _
                          strFileName    ,  _
                          strNewFileName    )

        Case CONST_TAKEOWNERSHIP

            Call TakeOwnOfFile(strServer      ,  _
                               strUserName    ,  _
                               strPassword    ,  _
                               objService     ,  _
                               objOutputFile  ,  _
                               strFileName       )
 
        Case CONST_ERROR
            'Do nothing.

        Case Else                    'Default -- should never happen
            Call Wscript.Echo("Error occurred in passing parameters.")

    End Select

    If NOT IsEmpty(objOutputFile) Then
        objOutputFile.Close
        Wscript.Echo "Results are saved in file " & strOutputFile & "."
    End If

End Sub
'********************************************************************
'*
'* Sub DeleteFile()
'*
'* Purpose: Delete a file.
'*
'* Input:   strServer           a machine name
'*          strUserName         the current user's name
'*          strPassword         the current user's password
'*          objService          The Wbem service
'*          objOutputFile       The Output File Object
'*          strFileName         The file to delete.
'*          blnForce            Overwrite Read-Only files
'*
'* Output:  Results are either printed on screen or saved in strOutputFile.
'*
'********************************************************************
Private Sub DeleteFile(ByVal strServer     ,  _
                       ByVal strUserName   ,  _
                       ByVal strPassword   ,  _
                       ByVal objService    ,  _
                       ByVal objOutputFile ,  _
                       ByVal strFileName   ,  _
                       ByVal blnForce         )

    ON ERROR RESUME NEXT

    Dim strWBEMClass
    Dim objFileSet, objInst
    Dim intFileCount

    strWBEMClass = "CIM_DataFile"

    strFileName = strDoubleBackSlashes(strFileName)

    Set objFileSet = objService.ExecQuery("SELECT * FROM " & strWBEMClass & _
                     " WHERE Name = """ & strFileName & """",,0)
    if objFileSet.Count = 0 then
         Call WriteLine("File not found.", objOutputFile)
         Exit Sub
    End If

    If blnErrorOccurred("Could not obtain " & strWBEMClass & " instance.") Then
        Exit Sub
    End If

    intFileCount = 0
    For Each objInst In objFileSet
        If blnCheckFile(objInst.Name,objService,blnForce) Then
            Call objInst.Delete()
            If NOT blnErrorOccurred("Could not delete " & objInst.Name) Then
                intFileCount = intFileCount + 1
            End If
        Else
            Call WriteLine("File is either marked as a system file, " _
                        & "or is read only.", objOutputFile)
            Call WriteLine("Use the /F switch to override.", objOutputFile)
        End If
    Next

    If blnErrorOccurred("Could not delete " & strFileName & ".") Then
        Exit Sub
    Else
        Call WriteLine("File deleted.", objOutputFile)
    End If

End Sub


'********************************************************************
'*
'* Sub RenameFile()
'*
'* Purpose: Rename a file.
'*
'* Input:   strServer           A machine name
'*          strUserName         The current user's name
'*          strPassword         The current user's password
'*          objService          The Wbem service
'*          objOutputFile       The Output File Object
'*          strFileName         The file to delete
'*          strNewFileName      The new file name
'*
'* Output:  Results are either printed on screen or saved in strOutputFile.
'*
'********************************************************************
Private Sub RenameFile(ByVal strServer      ,  _
                       ByVal strUserName    ,  _
                       ByVal strPassword    ,  _
                       ByVal objService     ,  _
                       ByVal objOutputFile  ,  _
                       ByVal strFileName    ,  _
                       ByVal strNewFileName    )


    ON ERROR RESUME NEXT

    Dim strWBEMClass
    Dim objInst, objFileSet

    strWBEMClass = "CIM_DataFile"

    strFileName = strDoubleBackSlashes(strFileName)

    Set objFileSet = objService.ExecQuery("SELECT * FROM " & strWBEMClass _
                     & " WHERE Name = """ & strFileName & """",,0)

    if objFileSet.Count = 0 then
        Call Writeline ("File not found.", objOutputFile)
        Exit Sub
    End If

    If blnErrorOccurred("Could not obtain " & strWBEMClass & " instance.") Then
        Exit Sub
    End If

    For Each objInst in objFileSet
        Call objInst.Rename(strNewFileName)
            If blnErrorOccurred("Could not copy " & strFileName & ".") Then
                Exit Sub
            Else
                Call WriteLine("File renamed.", objOutputFile)
           End If
    Next

End Sub


'********************************************************************
'*
'* Sub CopyFile()
'*
'* Purpose: Copy a file.
'*
'* Input:   strServer           a machine name
'*          strUserName         the current user's name
'*          strPassword         the current user's password
'*          objService          The Wbem service
'*          objOutputFile       The Output File Object
'*          strFileName         The file to copy.
'*          strNewFileName      The new file name.
'*
'* Output:  Results are either printed on screen or saved in strOutputFile.
'*
'********************************************************************
Private Sub CopyFile(ByVal strServer      ,  _
                     ByVal strUserName    ,  _
                     ByVal strPassword    ,  _
                     ByVal objService     ,  _
                     ByVal objOutputFile  ,  _
                     ByVal strFileName    ,  _
                     ByVal strNewFileName    )

    ON ERROR RESUME NEXT

    Dim strWBEMClass
    Dim objInst, objFileSet

    strWBEMClass = "CIM_DataFile"

    strFileName = strDoubleBackSlashes(strFileName)

    Set objFileSet = objService.ExecQuery("SELECT * FROM " & _
                strWBEMClass & " WHERE Name = """ & strFileName & """",,0)

    if objFileSet.Count = 0 then
        Call Writeline ("File not found.", objOutputFile)
        Exit Sub
    End If

    If blnErrorOccurred("Could not obtain " & strWBEMClass & " instance.") Then
        Exit Sub
    End If

    For Each objInst in objFileSet
        Call objInst.Copy(strNewFileName)
            If blnErrorOccurred("Could not copy " & strFileName & ".") Then
                Exit Sub
            Else
                Call WriteLine("File copied.", objOutputFile)
           End If
    Next

End Sub

'********************************************************************
'* Sub TakeOwnOfFile()
'*
'* Purpose: Take ownership of a file.
'*
'* Input:   strServer           a machine name
'*          strUserName         the current user's name
'*          strPassword         the current user's password
'*          objService          The Wbem service
'*          objOutputFile       The Output File Object
'*          strFileName         the file to sieze
'*
'* Output:  Results are either printed on screen or saved in strOutputFile.
'*
'********************************************************************
Private Sub TakeOwnOfFile(ByVal strServer      ,  _
                          ByVal strUserName    ,  _
                          ByVal strPassword    ,  _
                          ByVal objService     ,  _
                          ByVal objOutputFile  ,  _
                          ByVal strFileName       )

    ON ERROR RESUME NEXT

    Dim objFileSystem, objInst
    Dim strWBEMClass

    strWBEMClass = "CIM_DataFile"

    strFileName = strDoubleBackSlashes(strFileName)

    Set objInst = objService.Get(strWBEMClass & "=" & """" & strFileName & """")
    If blnErrorOccurred("Could not obtain " & strWBEMClass & " instance.") Then
        Exit Sub
    End If

    Call objInst.TakeOwnerShip()
    If blnErrorOccurred("Could not take ownership of " & strFileName & ".") Then
        Exit Sub
    Else
        Call WriteLine("Ownership taken.", objOutputFile)
    End If


End Sub

'********************************************************************
'*
'* Function intParseCmdLine()
'*
'* Purpose: Parses the command line.
'* Input:   
'*
'* Output:  strServer          a remote server ("" = local server")
'*          strUserName        the current user's name
'*          strPassword        the current user's password
'*          strOutputFile      an output file name
'*          strTaskCommand     one of /list, /start, /stop /install /remove
'*                                    /dependents
'*          strDriverName      name of the DEVICE
'*          strStartMode       start mode of the DEVICE
'*          strDisplayName     Display name for the DEVICE.
'*          blnDetails         Extra information to be displayed on the output

'*
'********************************************************************
Private Function intParseCmdLine( ByRef strServer      ,  _
                                  ByRef strUserName    ,  _
                                  ByRef strPassword    ,  _
                                  ByRef strOutputFile  ,  _
                                  ByRef strTaskCommand ,  _
                                  ByRef strFileName    ,  _
                                  ByRef strNewFileName ,  _
                                  ByRef blnForce          )


    ON ERROR RESUME NEXT

    Dim strFlag
    Dim intState, intArgIter
    Dim objFileSystem

    If Wscript.Arguments.Count > 0 Then
        strFlag = Wscript.arguments.Item(0)
    End If

    If IsEmpty(strFlag) Then                'No arguments have been received
        intParseCmdLine = CONST_ERROR
        Wscript.Echo("Arguments are required")
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
        Select Case LCase(Wscript.arguments.Item(intArgIter))
  
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

            Case "/d"
                If Not blnGetArg ("Target File", strFileName, intArgIter) Then
                    intParseCmdLine = CONST_ERROR
                    Exit Function
                End If
                strTaskCommand = CONST_DELETE
                intParseCmdLine = CONST_PROCEED
                intArgIter = intArgIter + 1
               
            Case "/t"
                If Not blnGetArg ("Target File", strFileName, intArgIter) Then
                    intParseCmdLine = CONST_ERROR
                    Exit Function
                End If
                strTaskCommand = CONST_TAKEOWNERSHIP
                intParseCmdLine = CONST_PROCEED
                intArgIter = intArgIter + 1

            Case "/r"
                If Not blnGetArg ("Target File", strFileName, intArgIter) Then
                    intParseCmdLine = CONST_ERROR
                    Exit Function
                End If
                strTaskCommand = CONST_RENAME
                intParseCmdLine = CONST_PROCEED
                intArgIter = intArgIter + 1
              
            Case "/c"
                If Not blnGetArg ("Target File", strFileName, intArgIter) Then
                    intParseCmdLine = CONST_ERROR
                    Exit Function
                End If
                strTaskCommand = CONST_COPY
                intParseCmdLine = CONST_PROCEED
                intArgIter = intArgIter + 1

            Case "/n"
                If Not blnGetArg ("New File", strNewFileName, intArgIter) Then
                    intParseCmdLine = CONST_ERROR
                    Exit Function
                End If
                intArgIter = intArgIter + 1 

            Case Else 'We shouldn't get here
                Call Wscript.Echo("Invalid or misplaced parameter: " _
                   & Wscript.arguments.Item(intArgIter) & vbCRLF _
                   & "Please check the input and try again," & vbCRLF _
                   & "or invoke with '/?' for help with the syntax.")
                Wscript.Quit

        End Select

    Loop '** intArgIter <= Wscript.arguments.Count - 1

    If IsEmpty(intParseCmdLine) Then 
        intParseCmdLine = CONST_ERROR
        Wscript.Echo ("Required arguments [/D | /T | /C | /R] are missing.")
    End If
    If strTaskCommand = CONST_COPY then
        if isEmpty(strNewFileName) then
            intParseCmdLine = CONST_ERROR
            Wscript.Echo ("The /N parameter is required when copying.")
        End If
    End If
    If strTaskCommand = CONST_RENAME then
        if isEmpty(strNewFileName) then
            intParseCmdLine = CONST_ERROR
            Wscript.Echo ("The /N parameter is required when renaming a file.")
        End If
    End If
    
    

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
    Wscript.Echo "Perform various simple operations on a file."
    Wscript.Echo ""
    Wscript.Echo "SYNTAX:"
    Wscript.Echo "1.  Fileman.vbs /D <targetfile> | /T <targetfile>" 
    Wscript.Echo "               [/S <server>][/U <username>][/W <password>]"
    Wscript.Echo "               [/O <outputfile>]"
    Wscript.Echo ""
    Wscript.Echo "2.  Fileman.vbs /R <targetfile> | /C <targetfile>"
    Wscript.Echo "                /N <newfile>"
    Wscript.Echo "               [/S <server>][/U <username>][/W <password>]"
    Wscript.Echo "               [/O <outputfile>]"
    Wscript.Echo ""
    Wscript.Echo "PARAMETER SPECIFIERS:"
    Wscript.Echo "   /D            Delete File"
    Wscript.Echo "   /T            Take Ownership of file"
    Wscript.Echo "   /R            Rename File"
    Wscript.Echo "   /C            Copy File"   
    Wscript.Echo "   targetfile    The target file."
    Wscript.Echo "   newfile       The destination filename."
    Wscript.Echo "   server        A machine name."
    Wscript.Echo "   username      The current user's name."
    Wscript.Echo "   password      Password of the current user."
    Wscript.Echo "   outputfile    The output file name."
    Wscript.Echo ""
    Wscript.Echo "EXAMPLE:"
    Wscript.Echo "1. cscript Fileman.vbs /d c:\test.txt /s MyMachine2"
    Wscript.Echo "   Deletes the file test.txt on the machine MyMachine2."
    Wscript.Echo "2. cscript Fileman.vbs /r c:\test.txt /n c:\new.txt"
    Wscript.Echo "   Renames the file test.txt to new.txt on the " _
               & "local machine."
    Wscript.Echo ""
    Wscript.Echo "NOTE:"
    Wscript.Echo "1.   You must include the full path when specifying " _
               & "the files."
    Wscript.Echo ""
    Wscript.Echo "2.   Wildcards (*, ?) will not work."
End Sub
'********************************************************************
'* General Routines
'********************************************************************


'********************************************************************
'*
'* Function blnCheckFile()
'*
'* Purpose: Checks to see if file is read only or marked as a system file.
'*
'* Input:   strFileName           The file to check
'*          objService            The WBEM object
'*          blnForce              Whether to force an action on a read only file.
'*
'* Output:  Ture or False
'*
'********************************************************************

Function blnCheckFile(strFileName,objService,blnForce)


    ON ERROR RESUME NEXT

    Dim objFileSet
    Dim strWBEMClass
    Dim Inst

    If blnForce=True Then
        blnCheckFile=True
        Exit Function
    Else
        blnCheckFile=False
    End If
    strFileName = strDoubleBackSlashes(strFileName)

  
    strWBEMClass = "CIM_LogicalFile"

    Set objFileSet = objService.ExecQuery("SELECT * FROM " & _
                     strWBEMClass & " WHERE Name = """ & strFileName & """",,0)

    If blnErrorOccurred("Could not obtain " & strWBEMClass & " instance.") Then
        Exit Function
    End If

    For Each Inst In objFileSet
        If Inst.System = True then
            Exit Function
        ElseIf inst.Writeable = False then
            Exit Function
        Else
            blnCheckFile = True
        End If
    Next

End Function

'********************************************************************
'*  Function: strDoubleBackSlashes (strIn)
'*
'*  Purpose:  expand path string to use double node-delimiters;
'*            doubles ALL backslashes.
'*
'*  Input:    strIn    path to file or directory
'*  Output:            WMI query-food
'*
'*  eg:      c:\pagefile.sys     becomes   c:\\pagefile.sys
'*  but:     \\server\share\     becomes   \\\\server\\share\\
'*
'********************************************************************
Private Function strDoubleBackSlashes (strIn)
    Dim i, str, strC
    str = ""
    for i = 1 to len (strIn)
        strC = Mid (strIn, i, 1)
        str = str & strC
        if strC = "\" then str = str & strC
    next
    strDoubleBackSlashes = str
End Function


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
             "1. Using ""CScript FileMan.vbs arguments"" for Windows 95/98 or" _
             & vbCRLF & "2. Changing the default Windows Scripting Host " _
             & "setting to CScript" & vbCRLF & "    using ""CScript " _
             & "//H:CScript //S"" and running the script using" & vbCRLF & _
             "    ""FileMan.vbs arguments"" for Windows NT/2000." )
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


