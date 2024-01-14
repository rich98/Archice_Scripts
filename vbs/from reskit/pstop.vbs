'********************************************************************
'*
'* File:           PsTop.VBS
'* Created:        March 1999
'* Version:        1.0
'*
'*  Main Function:  Gets CPU information for a machine.
'*
'*  PsTop.VBS   [/S <server>] [/U <username>] [/W <password>] 
'*              [/O <outputfile>]
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

    'Make sure the host is csript, if not then abort
    VerifyHostIsCscript()

    'Parse the command line
    intOpMode = intParseCmdLine(strServer     ,  _
                                strUserName   ,  _
                                strPassword   ,  _
                                strOutputFile    )


Select Case intOpMode
  Case CONST_SHOW_USAGE
    Call ShowUsage()
  Case CONST_PROCEED
    Call ListJobs(strServer,        _
                  strOutputFile,    _
                  strUserName,      _
                  strPassword)

    Case CONST_ERROR
        'Do nothing.
    Case Else                    'Default -- should never happen
        Print "Error occurred in passing parameters."
End Select

'********************************************************************
'*
'* Sub      ListJobs()
'*
'* Purpose: Lists all jobs currently running on a machine.
'*
'* Input:   strServer           a machine name
'*          intWidth            the default column width
'*          strUserName         the current user's name
'*          strPassword         the current user's password
'*          strOutputFile       an output file name
'*
'* Output:  Results are either printed on screen or saved in strOutputFile.
'*
'********************************************************************
Private Sub ListJobs( strServer,        _
                      strOutputFile,    _
                      strUserName,      _
                      strPassword)

    ON ERROR RESUME NEXT

    Dim objFileSystem, objOutputFile, objService, objEnumerator, objInstance
    Dim strOutRow
	Dim objProcList(), lngTimeProp()
	Dim strOutColumn(8), intColumnWidths(8)
	Dim i, j, intProcIndex, intObjects, intStart, intFinish, intStep

    intColumnWidths(1) = 20
    intColumnWidths(2) = 10
    intColumnWidths(3) = 14
    
    ReDim strPropertyTypes(0)

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

    Set objEnumerator = objService.InstancesOf("Win32_Process")

	'Dimension the object array and the property array for 
	'    the number of objects
    Redim objProcList(CInt(objEnumerator.Count-1))
	Redim lngTimeProp(CInt(objEnumerator.Count-1))

    intObjects = 0
	'Fill the arrays as a prelude to sorting
    'Filter out processes for which no times are available
	For Each objInstance In objEnumerator

        If NOT IsNull(objInstance.UserModeTime) OR NOT _
          IsNull(objInstance.KernelModeTime) Then
            lngTimeProp(i) = 0
            If NOT IsNull(objInstance.UserModeTime) Then
                lngTimeProp(intObjects) = lngTimeProp(intObjects) + _
                    CLng(objInstance.UserModeTime)
            End If
            If NOT IsNull(objInstance.KernelModeTime) Then
                lngTimeProp(intObjects) = lngTimeProp(intObjects) + _
                   CLng(objInstance.KernelModeTime)
            End If
  	        Set objProcList(intObjects) = objInstance
            intObjects = intObjects + 1
        End If
	Next

    'Determine the number of processes to list
    If 15 < intObjects Then
	    intObjects = 15
	End If

    'Always sort in descending order
    Call SortArray(lngTimeProp, 0, objProcList)

    'Construct the header
    strOutColumn(1) = "Image Name"
    strOutColumn(2) = "PID"
    strOutColumn(3) = "CPU Time"

    For j = 1 To 3
        strOutColumn(j) = strPackString(strOutColumn(j), _
            intColumnWidths(j), 1, 1)
    Next 'j

    strOutRow = CStr("")
    For j = 1 To 3
         strOutRow = strOutRow & strOutColumn(j)
    Next 'j
    Call WriteLine(strOutRow, objOutputFile)	  

    'Output the data
    intStart = 0
    intFinish = intObjects - 1
    intStep = 1

    For i = intStart To intFinish Step intStep
    
        strOutColumn(1) = objProcList(i).Name
        strOutColumn(2) = CStr(CLng(objProcList(i).ProcessId))
        strOutColumn(3) = FormatTime(lngTimeProp(i))

        For j = 1 To 3
            strOutColumn(j) = strPackString(strOutColumn(j), _
                intColumnWidths(j), 1, 1)
        Next 'j

        strOutRow = ""
        For j = 1 To 3
            strOutRow = strOutRow & strOutColumn(j)
        Next 'j
        Call WriteLine(strOutRow, objOutputFile)	  
    Next 'i

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
                                  ByRef strOutputFile     )


    ON ERROR RESUME NEXT

    Dim strFlag
    Dim intState, intArgIter, intWidth
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
    Wscript.Echo "List processes according to cpu usage in descending order."
    Wscript.Echo ""
    Wscript.Echo "SYNTAX:"
    Wscript.Echo "  PsTop.vbs [/S <server>] [/U <username>]" _
                &" [/W <password>]"
    Wscript.Echo "  [/O <outputfile>]"
    Wscript.Echo ""
    Wscript.Echo "PARAMETER SPECIFIERS:"
    Wscript.Echo "   server        A machine name."
    Wscript.Echo "   username      The current user's name."
    Wscript.Echo "   password      Password of the current user."
    Wscript.Echo "   outputfile    The output file name."
    Wscript.Echo ""
    Wscript.Echo "EXAMPLE:"
    Wscript.Echo "1. cscript PsTop.vbs"
    Wscript.Echo "   List the jobs running on the current machine."
    Wscript.Echo "2. cscript PsTop.vbs /S MyMachine2"
    Wscript.Echo "   List the jobs running on the machine MyMachine2."

End Sub


'********************************************************************
'*
'* Sub SortArray()
'*
'* Purpose: Sorts two arrays based on the contents in one array.
'*
'* Input:   strArray    the array that contains the data to sort
'*          blnOrder    True for ascending and False for descending
'*          objArray2   an array that has exactly the same number of 
'*                      elements as strArray
'*                      and will be reordered with strArray
'*
'* Output:  The arrarys are returned in sort order.
'*
'* Note:    Repeating elements are not deleted.
'*
'********************************************************************
Private Sub SortArray(varSortData, blnAscend, objList)

    ON ERROR RESUME NEXT

    Dim i, j, intUbound
    Dim blnSwapped
    Dim objSave

    If IsArray(varSortData) Then
        intUbound = UBound(varSortData)
    Else
        Wscript.Echo("Argument is not an array!")
        Exit Sub
    End If

    'This is true if a swap occurs and false otherwise
    blnSwapped = False

    blnAscend = CBool(blnAscend)
    If Err.Number Then
        Wscript.Echo("Argument is not a boolean!")
        Exit Sub
    End If

    If blnAscend Then

        Do
            blnSwapped = False

            For i = 0 To intUbound - 1
                If varSortData(i) > varSortData(i+1) Then
                    Call Swap( varSortData(i),  varSortData(i+1) )
                    Set objSave       = objList(i+1)
                    Set objList(i+1)  = objList(i)
                    Set objList(i)    = objSave
                    blnSwapped = True
                End If
            Next 'i

        Loop While blnSwapped

    Else 'Descend
        Do
            blnSwapped = False

            For i = 0 To intUbound - 1
                If varSortData(i) < varSortData(i+1) Then
                    Call Swap( varSortData(i),  varSortData(i+1) )
                    Set objSave       = objList(i+1)
                    Set objList(i+1)  = objList(i)
                    Set objList(i)    = objSave
                    blnSwapped = True
                End If
            Next 'i

        Loop While blnSwapped
    End If

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
    strA    = strB
    strB    = strTemp

End Sub

'********************************************************************
'*
'* Function FormatTime()
'*
'* Purpose: Converts milliseconds to Hour:Min:Sec format.
'*
'* Input:   lngMillSecs - number of milliseconds
'*
'* Output:  Returns time elapsed in Hour:Min:Sec format.
'*
'********************************************************************
Private Function FormatTime(lngMillSecs)
    Dim lngHour, lngMin, lngSec
    Dim strVal

    lngHour = Int(lngMillSecs / (60 * 60 * 1000))
    lngMin  = Int(lngMillSecs / (60 * 1000) - lngHour * 60)
    lngSec  = Int(lngMillSecs / 1000 - (lngMin * 60 + (60 * 60) * lngHour))

    FormatTime = lngHour

    If lngMin < 10 Then
        strVal = "0" & lngMin
    Else
        strVal = CStr(lngMin)
    End If
    FormatTime = FormatTime & ":" & strVal

    If lngSec < 10 Then
        strVal = "0" & lngSec
    Else
        strVal = CStr(lngSec)
    End If
    FormatTime = FormatTime & ":" & strVal

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
             "1. Using ""CScript PSTop arguments"" for Windows 95/98 or" _
             & vbCRLF & "2. Changing the default Windows Scripting Host " _
             & "setting to CScript" & vbCRLF & "    using ""CScript " _
             & "//H:CScript //S"" and running the script using" & vbCRLF & _
             "    ""PSTop arguments"" for Windows NT/2000." )
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
