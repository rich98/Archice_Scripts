'********************************************************************
'*
'* File:           CodecFile.vbs
'* Created:        March 1999
'* Version:        1.0
'*
'*  Main Function:  Outputs Information on Codec Files.
'*
'*  CodecFile.vbs [/S <server>] [/U <username>] [/W <password>] 
'*                [/O <outputfile>] [/D]
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

    'Declare variables
    Dim intOpMode, i
    Dim strServer, strUserName, strPassword, strOutputFile
    Dim blnDetails 

    'Make sure the host is csript, if not then abort
    VerifyHostIsCscript()

    'Parse the command line
    intOpMode = intParseCmdLine(strServer     ,  _
                                strUserName   ,  _
                                strPassword   ,  _
                                strOutputFile ,  _
                                blnDetails       )

    Select Case intOpMode

        Case CONST_SHOW_USAGE
            Call ShowUsage()

        Case CONST_PROCEED                 
            Call CodecFiles(strServer     , _
                            strOutputFile , _
                            strUserName   , _
                            strPassword   , _
                            blnDetails      )

        Case CONST_ERROR
            'Do Nothing

        Case Else                    'Default -- should never happen
            Call Wscript.Echo("Error occurred in passing parameters.")

    End Select

'********************************************************************
'* End of Script
'********************************************************************

'********************************************************************
'*
'* Sub CodecFiles
'*
'* Purpose: Outputs Information on Codec Files.
'*
'* Input:   strServer           a machine name
'*          strOutputFile       an output file name
'*          strUserName         the current user's name
'*          strPassword         the current user's password
'*          blnDetails          Extra information to be displayed
'*
'* Output:  Results are either printed on screen or saved in strOutputFile.
'*
'********************************************************************
Private Sub CodecFiles(strServer  ,  _
                       OutputFile ,  _
                       strUserName,  _
                       strPassword,  _
                       blnDetails    )
                       
    ON ERROR RESUME NEXT


    Dim objFileSystem, objOutputFile, objService, objSet, obj
    Dim strLine, strClass
    Dim n
    Dim intWidth(16)

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

    strClass = "Win32_CodecFile"

    'Get the first instance
    Set objSet = objService.InstancesOf(strClass)
    If blnErrorOccurred ("obtaining the "& strClass) Then Exit Sub

    i = objSet.Count
    WriteLine vbCRLF & CStr(i) & " instance" & strIf(i<>1,"s","") & " of " _
        & strClass& " on " & strServer & strIf(i>0,":","."), objOutputFile

    for i = 0 to 15: intWidth(i) = 4: next 'default widths
    intWidth(5) = 10 'FileSize
    intWidth(8) = 7 'Status

    'find column widths, ignoring headers & using 2 spacers
    for i = 0 to 15: intWidth(i) = intWidth(i) - 2: next
    For Each obj In objSet
        n = Len (obj.Name        ): if intWidth( 0) < n then intWidth( 0) = n
        n = Len (obj.Group       ): if intWidth( 1) < n then intWidth( 1) = n
        n = Len (obj.FileType    ): if intWidth( 2) < n then intWidth( 2) = n
        n = Len (obj.Version     ): if intWidth( 4) < n then intWidth( 3) = n
        n = Len (obj.Manufacturer): if intWidth( 3) < n then intWidth( 4) = n
        n = Len (obj.FileSize) + 1: if intWidth( 5) < n then intWidth( 5) = n
        n = Len (strFormatMOFTime(obj.CreationDate))
        if intWidth( 6) < n then intWidth( 6) = n
        n = Len (obj.Description ): if intWidth( 7) < n then intWidth( 7) = n
        n = Len (obj.Status      ): if intWidth( 8) < n then intWidth( 8) = n
        n = Len (obj.CSName      ): if intWidth(15) < n then intWidth(15) = n
    Next
    for i = 0 to 15: intWidth(i) = intWidth(i) + 2: next
    'print header
    strLine = Empty
    strLine = strLine + strPackString ("Name"        ,intWidth( 0),1,1)
    strLine = strLine + strPackString ("Group"       ,intWidth( 1),1,1)
    strLine = strLine + strPackString ("FileType    ",intWidth( 2),1,1)
    strLine = strLine + strPackString ("Version"     ,intWidth( 3),1,1)
    If blnDetails then
        strLine = strLine + strPackString ("Manufacturer", _
                  intWidth( 4),1,1)
        strLine = strLine + strPackString ("FileSize "   , _
                  intWidth( 5),0,1)
        strLine = strLine + strPackString ("Description ", _
                  intWidth( 7),1,1)
        strLine = strLine + strPackString ("Status      ", _
                  intWidth( 8),1,1)
        strLine = strLine + strPackString ("Readable    ", _
                  intWidth( 9),1,1)
        strLine = strLine + strPackString ("Writeable   ", _
                  intWidth(10),1,1)
        strLine = strLine + strPackString ("System      ", _
                  intWidth(11),1,1)
        strLine = strLine + strPackString ("Archive     ", _
                  intWidth(12),1,1)
        strLine = strLine + strPackString ("Hidden      ", _
                  intWidth(13),1,1)
        strLine = strLine + strPackString ("Encrypted   ", _
                  intWidth(14),1,1)
        strLine = strLine + strPackString ("CSName      ", _
                  intWidth(15),1,1)
    End If

    WriteLine " ", objOutputFile
    WriteLine strLine, objOutputFile

    'print header line
    n=0: for i = 0 to strIf(blnDetails,15,3): n = n + intWidth(i): next
    WriteLine Replace (Space(n), " ", "-"), objOutputFile
    'print records for each instance of class
    For Each obj In objSet
        strLine = Empty
        strLine = strLine + strPackString (obj.Name         ,intWidth( 0),1,1)
        strLine = strLine + strPackString (obj.Group        ,intWidth( 1),1,1)
        strLine = strLine + strPackString (obj.FileType     ,intWidth( 2),1,1)
        strLine = strLine + strPackString (obj.Version      ,intWidth( 3),1,1)
        if blnDetails then
            strLine = strLine + strPackString (obj.Manufacturer ,intWidth( 4),1,1)
            strLine = strLine + strPackString (obj.FileSize &" ",intWidth( 5),0,1)
            strLine = strLine + strPackString (obj.Description  ,intWidth( 7),1,1)
            strLine = strLine + strPackString (obj.Status       ,intWidth( 8),1,1)
            strLine = strLine + strPackString (strYesOrNo(obj.Readable) ,intWidth( 9),1,1)
            strLine = strLine + strPackString (strYesOrNo(obj.Writeable),intWidth(10),1,1)
            strLine = strLine + strPackString (strYesOrNo(obj.System)   ,intWidth(11),1,1)
            strLine = strLine + strPackString (strYesOrNo(obj.Archive)  ,intWidth(12),1,1)
            strLine = strLine + strPackString (strYesOrNo(obj.Hidden)   ,intWidth(13),1,1)
            strLine = strLine + strPackString (strYesOrNo(obj.Encrypted),intWidth(14),1,1)
            strLine = strLine + strPackString (obj.CSName       ,intWidth(15),1,1)
        end if

        WriteLine strLine, objOutputFile
    Next

    If IsObject(objOutputFile) Then
        objOutputFile.Close
        Call Wscript.Echo ("Results are saved in file " & strOutputFile & ".")
    End If

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
Private Function intParseCmdLine( ByRef strServer,        _
                                  ByRef strUserName,      _
                                  ByRef strPassword,      _
                                  ByRef strOutputFile,    _
                                  ByRef blnDetails        )


    ON ERROR RESUME NEXT

    Dim strFlag
    Dim intState, intArgIter
    Dim objFileSystem

    If Wscript.Arguments.Count > 0 Then
        strFlag = Wscript.arguments.Item(0)
    End If

    If IsEmpty(strFlag) Then                'No arguments have been received
        intParseCmdLine = CONST_PROCEED
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

            Case "/d"
                blnDetails = True
                intArgIter = intArgIter + 1

            Case Else 'We shouldn't get here
                Call Wscript.Echo("Invalid or misplaced parameter: " _
                   & Wscript.arguments.Item(intArgIter) & vbCRLF _
                   & "Please check the input and try again," & vbCRLF _
                   & "or invoke with '/?' for help with the syntax.")
                Wscript.Quit

        End Select

    Loop '** intArgIter <= Wscript.arguments.Count - 1

    If IsEmpty(intParseCmdLine) Then _
        intParseCmdLine = CONST_PROCEED

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
    Wscript.Echo "Outputs Information on Codec Files."
    Wscript.Echo ""
    Wscript.Echo "SYNTAX:"
    Wscript.Echo "  CodecFile.vbs [/S <server>] [/U <username>]" _
                &" [/W <password>] [/D]"
    Wscript.Echo "  [/O <outputfile>]"
    Wscript.Echo ""
    Wscript.Echo "PARAMETER SPECIFIERS:"
    Wscript.Echo "   server        A machine name."
    Wscript.Echo "   username      The current user's name."
    Wscript.Echo "   password      Password of the current user."
    Wscript.Echo "   outputfile    The output file name."
    Wscript.Echo "   /D            Show Details (wide format)."
    Wscript.Echo ""
    Wscript.Echo "EXAMPLE:"
    Wscript.Echo "1. cscript CodecFile.vbs"
    Wscript.Echo "   Get the codec file information for the current machine."
    Wscript.Echo "1. cscript CodecFile.vbs /S MyMachine2"
    Wscript.Echo "   Get the codec file information for the machine MyMachine2."

End Sub

'********************************************************************
'* General Routines
'********************************************************************

'********************************************************************
'*
'* Function strYesOrNo(blnB)
'*
'* Purpose: To give a boolean value to a "Yes" or "No"
'*
'* Input:   blnB    A BooleanValue
'*
'* Output:  "Yes" if Ture and "No" if False 
'*
'********************************************************************
Private Function strYesOrNo(blnB)

    strYesOrNo = "No"
    If blnB Then strYesOrNo = "Yes"

End Function

'********************************************************************
'*
'* Function strIF (blnTest, strTrue, strFalse) 
'*
'* Purpose: To replace a Boolean Value with a stricg
'*
'* Input:   blnB    A BooleanValue
'*
'* Output:  strTrue or strFalse
'*
'********************************************************************
Private Function strIF (blnTest, strTrue, strFalse) 

    If blnTest Then strIF = strTrue Else strIF = strFalse End If

End Function

'********************************************************************
'*
'* Function strFormatMOFTime(strDate)
'*
'* Purpose: Formats the date in WBEM to a readable Date
'*
'* Input:   blnB    A WBEM Date
'*
'* Output:  a string 
'*
'********************************************************************

Private Function strFormatMOFTime(strDate)
	Dim str
	str = Mid(strDate,1,4) & "-" _
           & Mid(strDate,5,2) & "-" _
           & Mid(strDate,7,2) & ", " _
           & Mid(strDate,9,2) & ":" _
           & Mid(strDate,11,2) & ":" _
           & Mid(strDate,13,2)
	strFormatMOFTime = str
End Function

'********************************************************************
'*
'* Function strYesOrNo(blnB)
'*
'* Purpose: To give a boolean value to a string
'*
'* Input:   blnB    A BooleanValue
'*
'* Output:  "Yes" if Ture and "No" if False 
'*
'********************************************************************
Private Function strYesOrNo(blnB)
    strYesOrNo = "No"
    If blnB Then strYesOrNo = "Yes"
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
             "1. Using ""CScript CodecFile.vbs arguments"" for Windows 95/98 or" _
             & vbCRLF & "2. Changing the default Windows Scripting Host " _
             & "setting to CScript" & vbCRLF & "    using ""CScript " _
             & "//H:CScript //S"" and running the script using" & vbCRLF & _
             "    ""CodecFile.vbs arguments"" for Windows NT/2000." )
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
