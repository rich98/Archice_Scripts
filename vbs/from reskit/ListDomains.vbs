
'********************************************************************
'*
'* File:        LISTDOMAINS.VBS
'* Created:     August 1998
'* Version:     1.0
'*
'* Main Function: Lists all domains within a namespace.
'* Usage: LISTDOMAINS.VBS adspath [/O:outputfile] [/U:username] [/W:password] [/Q]
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
Dim intOpMode, i
Dim strADsPath, strUserName, strPassword, blnQuiet, strOutputFile
ReDim strArgumentArray(0)

'Initialize variables
strArgumentArray(0) = ""
strADsPath = ""
strUserName = ""
strPassword = ""
strOutputFile = ""
blnQuiet = False

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
            "1. Using ""CScript LISTDOMAINS.vbs arguments"" for Windows 95/98 or" & vbCRLF & _
            "2. Changing the default Windows Scripting Host setting to CScript" & vbCRLF & _
            "    using ""CScript //H:CScript //S"" and running the script using" & vbCRLF & _
            "    ""LISTDOMAINS.vbs arguments"" for Windows NT."
        WScript.Quit
    Case Else
        WScript.Quit
End Select

'Parse the command line
intOpMode = intParseCmdLine(strArgumentArray, strADsPath, blnQuiet, _
            strUserName, strPassword, strOutputFile)
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
        Call GetDomains(strADsPath, strUserName, strPassword, strOutputFile)
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
'* Output:  strADsPath          the ADsPath of the root
'*          strUserName         name of the current user
'*          strPassword         password of the current user
'*          blnQuiet            specifies whether to suppress messages
'*          strOutputFile       an output file name
'*          intParseCmdLine     is set to one of CONST_ERROR, CONST_SHOW_USAGE, CONST_PROCEED.
'*
'********************************************************************

Private Function intParseCmdLine(strArgumentArray, strADsPath, blnQuiet, _
    strUserName, strPassword, strOutputFile)

    ON ERROR RESUME NEXT

    Dim i, strFlag

    strFlag = strArgumentArray(0)

    If strFlag = "" Then                'No arguments have been received
        Print "Arguments are required."
        intParseCmdLine = CONST_ERROR
        Exit Function
    End If

    'Help is needed
    If (strFlag="help") OR (strFlag="/h") OR (strFlag="\h") OR (strFlag="-h") _
        OR (strFlag = "\?") OR (strFlag = "/?") OR (strFlag = "?") OR (strFlag="h") Then
        intParseCmdLine = CONST_SHOW_USAGE
        Exit Function
    End If

    strADsPath = strFlag        'The first parameter must be the ADsPath.

    For i = 1 to UBound(strArgumentArray)
        strFlag = Left(strArgumentArray(i), InStr(1, strArgumentArray(i), ":")-1)
        If Err.Number Then            'An error occurs if there is no : in the string
            Err.Clear
            Select Case LCase(strArgumentArray(i))
                Case "/q"
                    blnQuiet = True
                Case Else
                    Print "Invalid flag " & strArgumentArray(i) & "."
                    Print "Please check the input and try again."
                    intParseCmdLine = CONST_ERROR
                    Exit Function
            End Select
        Else
            Select Case LCase(strFlag)
                Case "/u"
                    strUserName = Right(strArgumentArray(i), Len(strArgumentArray(i))-3)
                Case "/w"
                    strPassword = Right(strArgumentArray(i), Len(strArgumentArray(i))-3)
                Case "/o"
                    strOutputFile = Right(strArgumentArray(i), Len(strArgumentArray(i))-3)
                Case Else
                    Print "Invalid flag " & strFlag & "."
                    Print "Please check the input and try again."
                    intParseCmdLine = CONST_ERROR
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

    Wscript.Echo ""
    Wscript.Echo "Lists all domains within a namespace." & vbCRLF
    Wscript.Echo "LISTDOMAINS.VBS adspath [/O:outputfile]"
    Wscript.Echo "[/U:username] [/W:password] [/Q]"
    Wscript.Echo "   /O, /U, /W    Parameter specifiers."
    Wscript.Echo "   adspath       The ADsPath of the namespace."
    Wscript.Echo "   outputfile    The output file name."
    Wscript.Echo "   username      Username of the current user."
    Wscript.Echo "   password      Password of the current user."
    Wscript.Echo "   /Q            Suppresses all output messages." & vbCRLF
    Wscript.Echo "EXAMPLES:"
    Wscript.Echo "1. LISTDOMAINS.VBS LDAP://dc=Foo,dc=com"
    Wscript.Echo "   lists all domains within ""dc=Foo,dc=com""."
    Wscript.Echo "2. LISTDOMAINS.VBS WinNT:"
    Wscript.Echo "   lists all domains under WinNT:."

End Sub

'********************************************************************
'*
'* Sub GetDomains()
'* Purpose: Lists all domains withing a namespace.
'* Input:   strADsPath      the ADsPath of the root
'*          strUserName     name of the current user
'*          strPassword     password of the current user
'*          blnQuiet        specifies whether to suppress messages
'*          strOutputFile   an output file name
'* Output:  The domain names are either printed on screen or saved in strOutputFile.
'*
'********************************************************************

Private Sub GetDomains(strADsPath, strUserName, strPassword, strOutputFile)

    ON ERROR RESUME NEXT

    Dim strProvider, objRoot, objDomain, objFile, strFileName, objOutputFile, i

    objOutputFile = ""

    If strOutputFile <> "" Then
        'Create a filesystem object
        set objFile = CreateObject("Scripting.FileSystemObject")
        'Open the file for output
        set objOutputFile = objFile.OpenTextFile(strOutputFile, 8, True)
        If Err.Number then
            Print "Error 0x" & CStr(Hex(Err.Number)) & " occurred in opening file " _
                & strOutputFile
            If Err.Description <> "" Then
                Print "Error description: " & Err.Description & "."
            End If
            Err.Clear
            Exit Sub
        End If
    End If

    strProvider = Left(strADsPath, InStr(1, strADsPath, ":")-1)
    If Err.Number then
        Print "The provider """ & strADsPath & """ is not supported."
        Print "Make sure add : at the end."
        Err.Clear
        Exit Sub
    End If
    Select Case strProvider
        Case "WinNT"
            Call GetDomainsWinNT(strADsPath, strUserName, strPassword, objOutputFile)
        Case "LDAP"
            Call GetDomainsLDAP(strADsPath, strUserName, strPassword, objOutputFile)
        Case Else
            Print "The provider """" & strProvider & """" is not supported."
            Exit Sub
    End Select

    If strOutputFile <> "" Then
        objOutputFile.Close
        Wscript.Echo "Results are saved in file " & strOutputFile & "."
    End If

End Sub

'********************************************************************
'*
'* Sub GetDomainsLDAP()
'* Purpose: Lists all domains under a root with LDAP provider.
'* Input:   strADsPath      the ADsPath of the root
'*          strUserName     name of the current user
'*          strPassword     password of the current user
'*          objOutputFile   an output file object
'* Output:  The domain names are either printed on screen or saved in objOutputFile.
'*
'********************************************************************

Private Sub GetDomainsLDAP(strADsPath, strUserName, strPassword, objOutputFile)

    ON ERROR RESUME NEXT

    Dim objConnect, objCommand, objRecordSet, intCount
    Dim strPathCopy, strCriteria, strProperties, strScope, k

    strPathCopy =  "<" & strADsPath & ">;"
    strCriteria = "(ObjectClass=Domain);"
    strProperties = "Name, ADsPath;"
    strScope = "SubTree"

    Set objConnect = CreateObject("ADODB.Connection")
    If Err.Number then
        Print "Error 0x" & CStr(Hex(Err.Number)) & " ocurred in opening a connection."
        If Err.Description <> "" Then
            Print "Error description: " & Err.Description & "."
        End If
        Exit Sub
    End If

    objConnect.Provider = "ADsDSOObject"

    If strUserName = "" then
        objConnect.Open "Active Directory Provider"
    Else
        objConnect.Open "Active Directory Provider", strUserName, strPassword
    End If

    If Err.Number then
        Print "Error 0x" & CStr(Hex(Err.Number)) & " ocurred opening a provider."
        If Err.Description <> "" Then
            Print "Error description: " & Err.Description & "."
        End If
        Exit Sub
    End If

    Set objCommand = CreateObject("ADODB.Command")
    If Err.Number then
        Print "Error 0x" & CStr(Hex(Err.Number)) & " ocurred in creating the command object."
        If Err.Description <> "" Then
            Print "Error description: " & Err.Description & "."
        End If
        Exit Sub
    End If

    Set objCommand.ActiveConnection = objConnect

    'Set the query string
    objCommand.CommandText  = strPathCopy & strCriteria & strProperties & strScope
    objCommand.Properties("Page Size") = 100000                    'reset search properties
    objCommand.Properties("Timeout") = 300000 'seconds
    'objCommand.Properties("SearchScope") = 2

    'Let the user know what is going on
    Print "Start query: " & objCommand.CommandText
    'Execute the query
    Set objRecordSet = objCommand.Execute
    If Err.Number then
        Print "Error 0x" & CStr(Hex(Err.Number)) & " ocurred during the query."
        If Err.Description <> "" Then
            Print "Error description: " & Err.Description & "."
        End If
        Print "Clear the error and continue."
        Err.Clear
    End If
    Print "Finished the query."

    'Get the total number of objects found.
    objRecordSet.MoveLast
    intCount = objRecordSet.RecordCount

    If intCount Then                'If intCount is not zero
        Print "Found " & intCount & " domains."
        objRecordSet.MoveFirst
        k = 0
        While Not objRecordSet.EOF
            k = k + 1
            WriteLine objRecordSet.Fields(0).Value, objOutputFile
            objRecordSet.MoveNext
        Wend
        Print "Results are saved in file """ & strFileName & """."
    Else
        Print "There is no domain within " & strADsPath & "."
    End If

End Sub

'********************************************************************
'*
'* Sub GetDomainsWinNT()
'* Purpose: Lists all domains under a root with WinNT provider.
'* Input:   strADsPath      the ADsPath of the root
'*          strUserName     name of the current user
'*          strPassword     password of the current user
'*          objOutputFile   an output file object
'* Output:  The domain names are either printed on screen or saved in objOutputFile.
'*
'********************************************************************

Private Sub GetDomainsWinNT(strADsPath, strUserName, strPassword, objOutputFile)

    ON ERROR RESUME NEXT

    Dim objRoot, objDomain, strProvider, objProvider, i

    Print "Looking for domains in " & strADsPath & "..."
    If strUserName = ""    then        'The current user is assumed
        set objRoot = GetObject(strADsPath)
    Else
        'Credentials are passed
        strProvider = Left(strADsPath, InStr(1, strADsPath, ":"))
        set objProvider = GetObject(strProvider)
        'Use user authentication
        set objRoot = objProvider.OpenDsObject(strADsPath,strUserName,strPassword,1)
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

    objRoot.Filter = Array("domain")
    i = 0
    For each objDomain in objRoot
        If Err.Number then
            Err.Clear
        Else
            i = i + 1
            WriteLine objDomain.Name, objOutputFile
        End If
    Next

    If i = 0 Then
        Print "There is no domain under " & strADsPath & "."
    Else
        Print "There are " & i & " domains under " & strADsPath & "."
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
'*                                                                  *
'*                           End of File                            *
'*                                                                  *
'********************************************************************

'********************************************************************
'*
'* Procedures calling sequence: LISTDOMAINS.VBS
'*
'*  intChkProgram
'*  intParseCmdLine
'*  ShowUsage
'*  GetDomains
'*      GetDomainsWinNT
'*          WriteLine
'*      GetDomainsLDAP
'*          WriteLine
'*
'********************************************************************
