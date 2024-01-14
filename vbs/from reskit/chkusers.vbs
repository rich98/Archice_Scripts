
'********************************************************************
'*
'* File:        CHKUSERS.VBS
'* Created:     August 1998
'* Version:     1.0
'*
'* Main Function: Checks a domain for users satisfying a given criteria.
'* Usage: CHKUSERS.VBS </A:adspath | /I:inputfile> [/P:properties] /C:criteria
'*        [/O:outputfile] [/U:username] [/W:password] [/Q] [/M] [/NQ]
'* Note: A brief description of how the code works can be found at the end of the file.
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
CONST CONST_FAILED                  = -2

'Declare variables
Dim strADsPath, strCriteria, strUserName, strPassword
Dim strInputFile, strOutputFile, blnMultiFiles, blnQuestion, intOpMode, i
ReDim strArgumentArray(0), strProperties(0), strPropertyValues(0)
ReDim strOperators(0), strPropertiesOut(0)

'Initialize variables
strADsPath = ""
strCriteria = ""
strUserName = ""
strPassword = ""
strInputFile = ""
strOutputFile = ""
blnMultiFiles = False
blnQuestion = True
strArgumentArray(0) = ""
strProperties(0) = ""
strPropertyValues(0) = ""
strOperators(0) = ""
strPropertiesOut(0) = "ADsPath"        'Default
intOpMode = 0
i = 0

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
            "1. Using ""CScript CHKUSERS.vbs arguments"" for Windows 95/98 or" & vbCRLF & _
            "2. Changing the default Windows Scripting Host setting to CScript" & vbCRLF & _
            "    using ""CScript //H:CScript //S"" and running the script using" & vbCRLF & _
            "    ""CHKUSERS.vbs arguments"" for Windows NT."
        WScript.Quit
    Case Else
        WScript.Quit
End Select

'Parse the command line
intOpMode = intParseCmdLine(strArgumentArray, strADsPath, strCriteria, blnMultiFiles, _
            strProperties, strPropertyValues, strOperators, strPropertiesOut, _
            strUserName, strPassword, strInputFile, strOutputFile, blnQuestion)
If Err.Number then
    Print "Error 0x" & CStr(Hex(Err.Number)) & " occurred in parsing the command line."
    If Err.Description <> "" Then
        Print "Error description: " & Err.Description & "."
    End If
    WScript.Quit
End If

Select Case intOpMode
    Case CONST_SHOW_USAGE
        call ShowUsage()
    Case CONST_PROCEED
        'First we need to conver the datatype of strPropertyValues.
        Call ConvertPropertyValues(strProperties, strPropertyValues)
        'Now we can call ChkUsers to do the rest of the job.
        Call ChkUsers(strADsPath, strCriteria, blnMultiFiles, strProperties, _
             strPropertyValues, strOperators, strPropertiesOut, strUserName, _
             strPassword, strInputFile, strOutputFile, blnQuestion)
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
'* Output:  strADsPath          ADsPath of a domain
'*          strCriteria         the search criteria with each comparison replaced by
'*                              a corresponding index
'*          blnMultiFiles       specifies whether to save results to multiple files
'*          strProperties       an array containing names of user properties to be checked
'*          strPropertyValues   an array containing a set of target values of user properties
'*          strOperators        a string array containing comparison operators,
'*                              including >, < and =
'*          strPropertiesOut    an array containing names of user properties to be retrieved
'*          strUserName         name of the current user
'*          strPassword         password of the current user
'*          strInputFile        an input file name
'*          strOutputFile       an output file name
'*          blnQuestion         specifies whether to use message box to get info
'*          intParseCmdLine     is set to one of CONST_ERROR, CONST_SHOW_USAGE, CONST_PROCEED.
'*
'********************************************************************

Private Function intParseCmdLine(strArgumentArray, strADsPath, strCriteria, blnMultiFiles, _
    strProperties, strPropertyValues, strOperators, strPropertiesOut, _
    strUserName, strPassword, strInputFile, strOutputFile, blnQuestion)

    ON ERROR RESUME NEXT

    Dim i, j, k, intUBound, strFlag

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

    'Get strADsPath, strUserName, strPassword, strOutputFile from the input.
    intUBound = UBound(strArgumentArray)
    For i = 0 To intUBound
        strFlag = LCase(Left(strArgumentArray(i), InStr(1, strArgumentArray(i), ":")-1))
        If Err.Number Then            'An error occurs if there is no : in the string
            Err.Clear
            Select Case LCase(strArgumentArray(i))
                Case "/m"
                    blnMultiFiles = True
                Case "/nq"
                    blnQuestion = False     'No input box. Answer Yes to it.
                Case Else
                    Print "Invalid flag " & strArgumentArray(i) & "."
                    Print "Please check the input and try again."
                    intParseCmdLine = CONST_ERROR
                    Exit Function
            End Select
        Else
            Select Case strFlag
                Case "/a"
                    strADsPath = FormatProvider(Right(strArgumentArray(i), Len(strArgumentArray(i))-3))
                Case "/p"
                    j = 0
                    strArgumentArray(i) = Right(strArgumentArray(i), _
                        Len(strArgumentArray(i))-3)
                    Do
                        k = InStr(1, strArgumentArray(i), ";")
                        If k Then
                            ReDim Preserve strPropertiesOut(j)
                            strPropertiesOut(j) = Trim(Left(strArgumentArray(i), k-1))
                            strArgumentArray(i) = Trim(Right(strArgumentArray(i), _
                                Len(strArgumentArray(i))-k))
                            j = j + 1
                        End If
                    Loop Until k = 0
                    ReDim Preserve strPropertiesOut(j)
                    strPropertiesOut(j) = strArgumentArray(i)
                Case "/i"
                    strInputFile = Right(strArgumentArray(i), Len(strArgumentArray(i))-3)
                Case "/o"
                    strOutputFile = Right(strArgumentArray(i), Len(strArgumentArray(i))-3)
                Case "/u"
                    strUserName = Right(strArgumentArray(i), Len(strArgumentArray(i))-3)
                Case "/w"
                    strPassword = Right(strArgumentArray(i), Len(strArgumentArray(i))-3)
                Case "/c"
                    'Preserve the criteria for later treatment.
                    strCriteria = Right(strArgumentArray(i), Len(strArgumentArray(i))-3)
                Case Else
                    Print "Invalid flag " & strFlag & "."
                    Print "Please check the input and try again."
                    intParseCmdLine = CONST_ERROR
                    Exit Function
            End Select
        End If
    Next

    intParseCmdLine = CONST_PROCEED

    'Check whether strCriteria is empty.
    If Trim(strCriteria) = "" Then
        Print "Please enter a criteria."
        intParseCmdLine = CONST_ERROR
        Exit Function
    Else
        'Get strProperties, strPropertyValues, strOperators from the criteria.
        If blnParseCriteria(strCriteria, strProperties, _
            strPropertyValues, strOperators) Then
            Print "An error occurred in parsing the criteria."
            Print "Please check the syntax and try again."
            intParseCmdLine = CONST_ERROR
            Exit Function
        End If
        'Check whether strCriteria is empty now.
        If Trim(strCriteria) = "" Then
            Print "Please enter a criteria."
            intParseCmdLine = CONST_ERROR
        Else
            'Check whether the syntx is correct.
            i = intEvalCriteria(strCriteria)
            If i = CONST_FAILED Then
                Print "Please check the syntax and try again."
                intParseCmdLine = CONST_ERROR
            End If
        End If
    End If

    'The ADsPath is required.
    If strADsPath = "" and strInputFile = "" Then
        Print "Please enter either an ADsPath or a file name."
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

Sub ShowUsage()

    Wscript.Echo ""
    Wscript.Echo "Checks a domain for users satisfying a given criteria." & vbCRLF
    Wscript.Echo "CHKUSERS.VBS </A:adspath | /I:inputfile> [/P:properties] /C:criteria "
    Wscript.Echo "             [/O:outputfile] [/U:username] [/W:password] [/M] [/NQ]"
    Wscript.echo "              /A, /I, /P, /C, /U, /W, /O" & vbCRLF
    Wscript.Echo "Parameter specifiers:"
    Wscript.echo "   adspath       ADsPath of a user object container."
    Wscript.Echo "   inputfile     A file containing ADsPaths of domains."
    Wscript.Echo "                 It can be used to check many domains at once."
    Wscript.Echo "   properties    Properties to be retrieved."
    Wscript.Echo "   criteria      Specifies what kind of users to look for."
    Wscript.Echo "   outputfile    The output file name."
    Wscript.echo "   username      Username of the current user."
    Wscript.echo "   password      Password of the current user."
    Wscript.Echo "   /M            Specifies that the output is written to multiple"
    Wscript.Echo "                 files to be created in the script."
    Wscript.Echo "   /NQ           Specifies files named Users* under the current"
    Wscript.Echo "                 directory be deleted without poping up a MsgBox."
    Wscript.Echo ""
    Wscript.Echo "EXAMPLE:"
    Wscript.Echo "   CHKUSERS.VBS /A:WinNT://FooFoo /P:FullName;Description"
    Wscript.Echo "   /C:""((LastLogin:>4/3/98 or LastLogin:<8/4/98)" _
               & " and AccountDisabled:=False)""" & vbCRLF
    Wscript.Echo "   gets the FullName and Description of all active users whose"
    Wscript.Echo "   last login is between 4/10/98 and 8/4/98." & vbCRLF
    Wscript.Echo "NOTES:"
    Wscript.Echo "1. The property name and the operator in the criteria must be"
    Wscript.Echo "   separated by a colon."
    Wscript.Echo "2. The criteria and any string including spaces must be "
    Wscript.Echo "   enclosed in quotes."
    Wscript.Echo "3. Any string within the criteria including spaces must be"
    Wscript.Echo "   enclosed in single quotes."

End Sub

'********************************************************************
'*
'* Function blnParseCriteria()
'* Purpose: Gets strProperties, strPropertyValues, strOperators from the criteria.
'* Input:   strCriteria         the search criteria
'* Output:  strCriteria         the search criteria with each comparison replaced by
'*                              a corresponding index
'*          strProperties       an array containing names of user properties to be checked
'*          strPropertyValues   an array containing a set of target values of user properties
'*          strOperators        a string array containing comparison operators,
'*                              including >, < and =
'*
'********************************************************************

Private Function blnParseCriteria(strCriteria, strProperties, _
    strPropertyValues, strOperators)

    ON ERROR RESUME NEXT

    Dim i, j, intColon, intQuote, intSpace, intBracket, strLeft, strRight, strTemp

    blnParseCriteria = False     'No error.
    strTemp = ""
    j = 0

    If strCriteria = "" Then
        Print "Please enter a criteria."
        blnParseCriteria = True            'An error.
        Exit Function
    End If
    intColon = InStr(1, strCriteria, ":")
    'Replace each comparison(including a property name, a value, and an operator)
    'with value of j and read property name, value and operators into corresponding arrays.
    Do While intColon    'If there is a : in the criteria
        If intColon = 1 Then    'This must be an error
            blnParseCriteria = True
            Exit Function
        End If
        ReDim Preserve strProperties(j), strPropertyValues(j), strOperators(j)
        strLeft = Trim(Left(strCriteria, intColon-1))
        strRight = Trim(Right(strCriteria, Len(strCriteria)-intColon))
        If strLeft = "" or strRight = "" Then
            Print "A property name or property value is missing."
            blnParseCriteria = True
            Exit Function
        End If

        'Now treat the left side.
        intBracket = InStrRev(strLeft, "(")
        intSpace = InStrRev(strLeft, " ")        'The first appearance of a space.
        If intSpace Then
            If intBracket and intBracket > intSpace    Then
                'Then strProperties(j) is down to the bracket.
                strProperties(j) = Trim(Right(strLeft, Len(strLeft)-intBracket))
                strTemp = strTemp & Left(strLeft, intBracket) & j & " "
            Else     'strProperties(j) is down to the space.
                strProperties(j) = Right(strLeft, Len(strLeft)-intSpace)
                strTemp = strTemp & Left(strLeft, intSpace) & j & " "
            End If
        Else    'If there is no space in strLeft
            If intBracket Then
                strProperties(j) = Trim(Right(strLeft, Len(strLeft)-intBracket))
                strTemp = strTemp & Left(strLeft, intBracket) & j & " "
            Else        'There is no space nor bracket.
                strProperties(j) = strLeft
                strTemp = strTemp & j & " "
            End If
        End If

        'Now treat the right side
        intQuote = InStr(strRight, "'")        'The first appearance of '.
        intSpace = InStr(strRight, " ")        'The first appearance of a space.
        intBracket = InStr(strRight, ")")        'The first appearance of a ).
        If intSpace Then    'If there is a space in strRight
            'If there is a ' in the left most part of strRight then
            'strPropertyValues(j) should be up to the next '.
            If intQuote and intSpace > intQuote Then
                'Get the position of the next '.
                intQuote = InStr(intQuote+1, strRight, "'")
                'It is an error to have a bracket between two single quotes.
                If intBracket and intQuote > intBracket Then
                    Print "A bracket is misplaced."
                    blnParseCriteria = True
                    Exit Function
                End If
                strPropertyValues(j) = Trim(Left(strRight, intQuote-1))
                'Get rid of the first '.
                strPropertyValues(j) = Replace(strPropertyValues(j), "'", "")
                strCriteria = Trim(Right(strRight, Len(strRight)-intQuote))
            Else
                'If the left most string ends up with a bracket,
                'strPropertyValues(j) should be up to the bracket.
                If intBracket and intSpace > intBracket Then
                    strPropertyValues(j) = Trim(Left(strRight, intBracket-1))
                    strCriteria = Trim(Right(strRight, Len(strRight)-intBracket+1))
                Else        'strPropertyValues(j) should be up to the space.
                    strPropertyValues(j) = Left(strRight, intSpace-1)
                    strCriteria = Trim(Right(strRight, Len(strRight)-intSpace+1))
                End If
            End If
        Else    'If there is no space in strRight
            If intQuote Then
                Print "A single quote is misplaced."
                blnParseCriteria = True
                Exit Function
            End If
            If intBracket Then
                strPropertyValues(j) = Trim(Left(strRight, intBracket-1))
                strCriteria = Trim(Right(strRight, Len(strRight)-intBracket+1))
            Else    'If there is no bracket then strPropertyValues(j) should be up to the end.
                strPropertyValues(j) = strRight
                strCriteria = ""
            End If
        End If

        'Now take care of the operator in strPropertyValues
        Select Case LCase(Left(strPropertyValues(j),1))
            Case ">"
                strOperators(j) = ">"
                strPropertyValues(j) = Right(strPropertyValues(j), Len(strPropertyValues(j))-1)
            Case "<"
                strOperators(j) = "<"
                strPropertyValues(j) = Right(strPropertyValues(j), Len(strPropertyValues(j))-1)
            Case "="
                strOperators(j) = "="
                strPropertyValues(j) = Right(strPropertyValues(j), Len(strPropertyValues(j))-1)
            Case Else    'Assume that an operator has been omitted.
                strOperators(j) = "="
        End Select

        If strPropertyValues(j) = "" Then
            Print "Warning: no value is entered for property """ & strProperties(j) & """."
        End If
        j = j + 1
        If strCriteria <> "" Then
            intColon = InStr(1, strCriteria, ":")
        Else
            intColon = 0
        End If
    Loop
    strCriteria = strTemp & strCriteria

End Function

'********************************************************************
'*
'* Sub ConvertPropertyValues()
'* Purpose: Converts elements of strPropertyValues to the right datatype.
'* Input:   strProperties       an array holding names of user properties
'*          strPropertyValues   an array holding the corresponding values of user properties
'* Output:  Elements of strPropertyValues are converted to the appropriate datatypes.
'*
'********************************************************************

Private Sub ConvertPropertyValues(strProperties, strPropertyValues)

    ON ERROR RESUME NEXT

    Dim i, strProperty

    For i = 0 To UBound(strProperties)
        strProperty = LCase(strProperties(i))
        If strProperty="badpasswordattempts" or strProperty="maxlogins" or _
            strProperty="maxstorage" or strProperty="maxpasswordage" or _
            strProperty="minpasswordage" or strProperty="passwordhistorylength" or _
            strProperty="userflags" or strProperty="codepage" or strProperty="countrycode" or _
            strProperty="primarygroupid" or strProperty="samaccounttype" Then
            strPropertyValues(i) = CLng(strPropertyValues(i))
        ElseIf strProperty="lastlogin" or strProperty="lastlogoff" or _
            strProperty="accountexpirationdate" Then
            strPropertyValues(i) = CDate(strPropertyValues(i))
        ElseIf strProperty="passwordneverexpires" or strProperty="usercannotchangepassword" _
            or strProperty = "accountdisabled" Then
            strPropertyValues(i) = CBool(strPropertyValues(i))
        ElseIf strProperty="passwordexpired" Then
            strPropertyValues(i) = CLng(-CBool(strPropertyValues(i)))
        End If
    Next
    If Err.Number Then
        Err.Clear
        Print "Please check the input datatype and try again."
        Wscript.Quit
    End If

End Sub

'********************************************************************
'*
'* Sub ChkUsers()
'* Purpose: Checks a domain for users against given criteria.
'* Input:   strADsPath          ADsPath of a domain
'*          strCriteria         the search criteria with each comparison replaced by
'*                              a corresponding index
'*          blnMultiFiles       specifies whether to save results to multiple files
'*          strProperties       an array containing names of user properties to be checked
'*          strPropertyValues   an array containing a set of target values of user properties
'*          strOperators        a string array containing comparison operators,
'*                              including >, < and =
'*          strPropertiesOut    an array containing names of user properties to be retrieved
'*          strUserName         name of the current user
'*          strPassword         password of the current user
'*          strInputFile        an input file name
'*          strOutputFile       an output file name
'*          blnQuestion         specifies whether to use message box to get info
'* Output:  Specified properties of users satisfying the criteria are either printed
'*          on screen or saved in file strOutputFile.
'*
'********************************************************************

Private Sub ChkUsers(strADsPath, strCriteria, blnMultiFiles, strProperties, _
    strPropertyValues, strOperators, strPropertiesOut, strUserName, strPassword, _
    strInputFile, strOutputFile, blnQuestion)

    ON ERROR RESUME NEXT

    Dim strProvider, objProvider, objDomain, objFileSystem, objInputFile, objOutputFile
    Dim intFound, intFiles, objFolder, colFiles, strMessage, i

    intFound = 0
    intFiles = 0

    'Check whether the Users series files exist in the current folder.
    'If they do, ask for permission to delete them.
    'The results will be saved into file Users* instead of strOutputFile.
    If blnMultiFiles Then
        'Create a filesystem object
        Set objFileSystem = CreateObject("Scripting.FileSystemObject")
        If Err.Number then
            Print "Error 0x" & CStr(Hex(Err.Number)) & " opening a filesystem object."
            If Err.Description <> "" Then
                Print "Error description: " & Err.Description & "."
            End If
            Exit Sub
        End If

        'get the current folder
        Set objFolder = objFileSystem.GetFolder(".")
        Set colFiles = objFolder.Files
        For Each objInputFile in colFiles
            'Check whether the file name starts with "users"
            If Left(LCase(objInputFile.name), 5) = "users" Then
                intFound = 1
                Exit For
            End If
        Next
        If intFound Then
            If blnQuestion Then
                strMessage = "All files named Users* in the current directory will be deleted."
                strMessage = strMessage & vbCRLF & "To save these files please move them"
                strMessage = strMessage & " to another directory before click the OK button."
                strMessage = strMessage & vbCRLF & "Click Cancel to quit the operation."
                'Ask the user for permission to delete files named Users*.
                i = MsgBox(strMessage, vbExclamation + vbOKCancel + vbDefaultButton2)
            Else
                i = vbCancel + 1        'Assign a value to i so it is not vbCancel
            End If
            If i = vbCancel Then
                Wscript.quit
            Else
                'Delete Users* files.
                For Each objInputFile in colFiles
                    If Left(LCase(objInputFile.name), 5) = "users" Then
                        objFileSystem.DeleteFile(objInputFile.name)        'Delete this file.
                    End If
                Next
            End If
        End If
        intFound = 0        'return it to zero.
        intFiles = 1        'initializes intFiles for next output file name.
        strOutputFile = "Users" & intFiles        'initializes the output file name.
    End If

    If strOutputFile = "" Then
        objFileSystem = ""
        objOutputFile = ""
    Else
        If Not IsObject(objFileSystem) Then
            'Create a filesystem object
            Set objFileSystem = CreateObject("Scripting.FileSystemObject")
            If Err.Number then
                Print "Error 0x" & CStr(Hex(Err.Number)) & " opening a filesystem object."
                If Err.Description <> "" Then
                    Print "Error description: " & Err.Description & "."
                End If
                Exit Sub
            End If
        End If
        'Open the file for output
        Set objOutputFile = objFileSystem.OpenTextFile(strOutputFile, 8, True)
        If Err.Number then
            Print "Error 0x" & CStr(Hex(Err.Number)) & " opening file " & strOutputFile
            If Err.Description <> "" Then
                Print "Error description: " & Err.Description & "."
            End If
            Exit Sub
        End If
    End If

    'Check the domain specified by /A:adspath.
    If strADsPath <> "" Then
        If strUserName = ""    then        'The current user is assumed
            Set objDomain = GetObject(strADsPath)
        Else                        'Credentials are passed
            strProvider = Left(strADsPath, InStr(1, strADsPath, ":"))
            Set objProvider = GetObject(strProvider)
            'Use user authentication
            Set objDomain = objProvider.OpenDsObject(strADsPath,strUserName,strPassword,1)
        End If
		If Err.Number then
			If CStr(Hex(Err.Number)) = "80070035" Then
				Print "Object " & strADsPath & " does not exist."
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
        'intFound is the number of users found not logged in for intDays days.
        intFound = intChkOneDomain(objDomain, strCriteria, strProperties,_
            strPropertyValues, strOperators, strPropertiesOut, objOutputFile)
        If blnMultiFiles Then        'close the output file
            objOutputFile.Close
            intFiles = intFiles + 1        'initializes intFiles for next output file name.
        End If
    End If

    'Check domains listed in /I:inputfile.
    If strInputFile <> "" Then
        If Not IsObject(objFileSystem) Then
            'Create a filesystem object
            Set objFileSystem = CreateObject("Scripting.FileSystemObject")
            If Err.Number then
                Print "Error 0x" & CStr(Hex(Err.Number)) & " opening a filesystem object."
                If Err.Description <> "" Then
                    Print "Error description: " & Err.Description & "."
                End If
                Exit Sub
            End If
        End If
        'Open the file for input
        Set objInputFile = objFileSystem.OpenTextFile(strInputFile)
        If Err.Number then
            Print "Error 0x" & CStr(Hex(Err.Number)) & " opening file " & strInputFile
            If Err.Description <> "" Then
                Print "Error description: " & Err.Description & "."
            End If
            Exit Sub
        End If
        'Read input file.
        While not objInputFile.AtEndOfStream
            'Get rid of leading and trailing spaces
            strADsPath = Trim(objInputFile.ReadLine)
            If strADsPath <> "" Then                        'Get rid of empty lines
                If strUserName = ""    then        'The current user is assumed
                    Set objDomain = GetObject(strADsPath)
                Else                        'Credentials are passed
                    strProvider = Left(strADsPath, InStr(1, strADsPath, ":"))
                    Set objProvider = GetObject(strProvider)
                    'Use user authentication
                    Set objDomain = objProvider.OpenDsObject(strADsPath,strUserName,_
                        strPassword,1)
                End If
				If Err.Number then
					If CStr(Hex(Err.Number)) = "80070035" Then
						Print "Object " & strADsPath & " does not exist."
					Else
						Print "Error 0x" & CStr(Hex(Err.Number)) & _
							" occurred in getting object " & strADsPath & "."
						If Err.Description <> "" Then
							Print "Error description: " & Err.Description & "."
						End If
					End If
					Err.Clear
					Exit Sub
				End If

                'Get the right file name
                If blnMultiFiles Then
                    'Change the output file name to "Users" & intFiles
                    strOutputFile = "Users" & intFiles
                    'Open the file for output
                    Set objOutputFile = objFileSystem.OpenTextFile(strOutputFile, 8, True)
                End If
                'intFound is the number of users found not logged in for intDays days.
                intFound = intFound + intChkOneDomain(objDomain, strCriteria, strProperties, _
                    strPropertyValues, strOperators, strPropertiesOut, objOutputFile)
                If blnMultiFiles Then
                    'Close the output file and initializes intFiles for next output file name.
                    objOutputFile.Close
                    intFiles = intFiles + 1
                End If
            End If
        Wend
        objInputFile.Close
    End If

    If blnMultiFiles Then
        If intFound Then
            strOutputFile = ""
            For i = 1 To intFiles-1
                strOutputFile = strOutputFile & "Users" & i & ", "
            Next
            'Get rid of the last two characters.
            strOutputFile = Left(strOutputFile, Len(strOutputFile)-2)
            Wscript.Echo  "Results are saved in files " & strOutputFile & "."
        End If
    Else
        If strOutputFile <> "" Then
            If intFound Then
                Wscript.Echo  "Results are saved in file " & strOutputFile & "."
            End If
            objOutputFile.Close
        End If
    End If

End Sub

'********************************************************************
'*
'* Sub intChkOneDomain()
'* Purpose:   Checks a domain for users against a given criteria.
'* Input:   objDomain           the domain to be checked
'*          strCriteria         the search criteria with each comparison replaced by
'*                              a corresponding index
'*          strProperties       an array containing names of user properties to be checked
'*          strPropertyValues   an array containing a set of target values of user properties
'*          strOperators        a string array containing comparison operators,
'*                              including >, < and =
'*          strPropertiesOut    an array containing names of user properties to be retrieved
'*          strOutputFile       an output file name
'* Output:  Specified properties of users satisfying the criteria are either printed
'*          on screen or saved in file strOutputFile. intChkOneDomain is set to the
'*          number of users found.
'*
'********************************************************************

Private Function intChkOneDomain(objDomain, strCriteria, strProperties, strPropertyValues,_
    strOperators, strPropertiesOut, objOutputFile)

    ON ERROR RESUME NEXT

    Dim i, intFound, intUBound, strPropertyValue, intResults(), strTemp, objADs

    intFound = 0
    intChkOneDomain = 0
    intUBound = UBound(strProperties)
    ReDim intResults(intUBound)
    objDomain.Filter = Array("user")

    For Each objADs in objDomain
        For i = 0 To intUBound
            'Get a user property.
            If blnGetOneProperty(objADs, strProperties(i), strPropertyValue) Then
                Print "Unable to get property " & strProperties(i)
                Exit Function
            End If

            'Compare the value with the criteria.
            intResults(i) = intCompare(strPropertyValue, strOperators(i), _
                strPropertyValues(i))
            If intResults(i) = CONST_FAILED Then
                Print "Failed to compare property " & strProperties(i) & " with " _
                    & strPropertyValues(i) & "."
                Exit Function
            End if
        Next

        'Copy criteria into strTemp so criteria is not modified by the subsequent operations.
        strTemp = strCriteria
        'Now replace the digits inserted in the criteria array with
        'the corresponding value from the above comparison.
        If blnCopyResults(intResults, strTemp) Then
            Print "Error occurred in copying an array."
            Err.Clear
            Exit Function
        End If

        'Evaluate the user properties to determine whether it satisfies the criteria.
        i = intEvalCriteria(strTemp)
        If i = CONST_FAILED Then
            Print "Failed to evaluate the expression."
            Exit Function
        ElseIf i Then        'If it satisfies the criteria.
            intFound = intFound + 1
            If blnPrintProperties(objADs, strPropertiesOut, objOutputFile) Then
                Print "Failed to get properties for user " & objADs.Name & "."
            End If
        End If
    Next

    Print intFound & " users satisfying the criteria have been found in " _
        & objDomain.ADsPath & "."
    intChkOneDomain = intFound

End Function

'********************************************************************
'*
'* Sub blnGetOneProperty()
'* Purpose: Gets one property of a given ADS object.
'* Input:   objADS              an ADS object
'*          strProperty         name of a property
'*          strPropertyValue    a string to save the value of the property
'* Output:  blnGetOneProperty is set to True if an error occurred and False otherwise.
'*
'********************************************************************

Function blnGetOneProperty(objADS, ByVal strProperty, ByRef strPropertyValue)

    ON ERROR RESUME NEXT

    Dim lngFlag, strResult, i, intUBound

    blnGetOneProperty = False
    strProperty = LCase(strProperty)

    Select Case strProperty
        Case "usercannotchangepassword"
            lngFlag = objADs.Get("UserFlags")
            If lngFlag = lngFlag and CONST_UF_PASSWORD_CAN_CHANGE    Then
                strPropertyValue = 0        'User can change password
            Else
                strPropertyValue = 1
            End If
        Case "passwordneverexpires"
            lngFlag = objADs.Get("UserFlags")
            If lngFlag = lngFlag or CONST_UF_DONT_EXPIRE_PASSWORD Then
                strPropertyValue = 1        'Password does not expire.
            Else
                strPropertyValue = 0
            End If
        Case Else
            strResult = objADS.Get(strProperty)
            If Err.Number Then
                Err.Clear            'The property is not available.
                If strProperty = "lastlogin" or strProperty = "lastlogoff" Then
                    strPropertyValue = CDate("1/1/1900")        'A date in the remote past.
                Else
                    blnGetOneProperty = True
                End If
            Else
                If IsArray(strResult) Then
                    Print strProperty & " is a multivalued property."
                    Print "The last value is used."
                    strPropertyValue = strResult(UBound(strResult))
                Else
                    strPropertyValue =  strResult
                End If
            End If
    End Select

End Function

'********************************************************************
'*
'* Function intCompare()
'* Purpose: Compares the value of a user property with the input value.
'* Input:   strValue1       the value of a user property
'*          strOperator     the comparison operator
'*          strValue2       the input value
'* Output:  intCompare = CONST_FAILED if an error occurred, otherwise
'*          it is 1 if the comparison evaluates to true and 0 if false.
'*
'********************************************************************

Private Function intCompare(strValue1, strOperator, strValue2)

    Dim i, strLeft1, strLeft2

    Select Case strOperator
        Case ">"
            If strValue1 > strValue2 Then
                intCompare = 1
            Else
                intCompare = 0
            End If
        Case "<"
            If strValue1 < strValue2 Then
                intCompare = 1
            Else
                intCompare = 0
            End If
        Case "="
            i = InStr(1, strValue2, "*")        'Check for wild card *
            If i > 1 Then
                strLeft1 = Left(strValue1, i-1)
                strLeft2 = Left(strValue2, i-1)
                If LCase(strLeft1) = LCase(strLeft2) Then
                    intCompare = 1
                Else
                    intCompare = 0
                End If
            ElseIf i = 1 Then        'As long as strValue1 is not empty, intCompare = 1.
                If strValue1 = "" Then
                    intCompare = 0
                Else
                    intCompare = 1
                End If
            Else
                If LCase(strValue1) = LCase(strValue2) Then
                    intCompare = 1
                Else
                    intCompare = 0
                End If
            End If
        Case Else
            Print "Operator " & strOperator & " is not supported."
            intCompare = CONST_FAILED
    End Select

End Function

'********************************************************************
'*
'* Function blnCopyResults()
'* Purpose: Replaces integers in strString with corresponding elements of intResults.
'* Input:   intResults      an array containing elements with a value of either 1 or 0
'*          strString       the original criteria string with each comparison unit
'*                          replaced by an integer
'* Output:  blnCopyResults = True if an error occurred and False otherwise.
'*
'********************************************************************

Private Function blnCopyResults(intResults, strString)

    Dim i, k, strLeft, strRight

    k = 0
    blnCopyResults = False        'No error.

    For i = 0 To UBound(intResults)
        k = k + 1
        'Start the search at position k and save the result in k.
        k = InStr(k, strString, CStr(i))
        strLeft = Left(strString, k-1)
        strRight = Right(strString, Len(strString)-k)
        If k Then
            strString = strLeft & intResults(i) & strRight
        Else
            blnCopyResults = True        'This is an error.
            Exit Function
        End If
    Next

End Function

'********************************************************************
'*
'* Function blnPrintProperties()
'* Purpose: Gets specified user properties and writes them either to a file or on screen.
'* Input:   objADS              an ADS object
'*          strPropertyArray    an array of properties
'*          objOutputFile       a file object for output
'* Output:  blnGetOneProperty is True if an error occurred and False otherwise.
'*
'********************************************************************

Function blnPrintProperties(objADS, strPropertyArray, objOutputFile)

    ON ERROR RESUME NEXT

    Dim i, strResult, strTemp, strOutput

    blnPrintProperties = False        'No error.
    strOutput = ""
    For i = 0 To UBound(strPropertyArray)
        Select Case LCase(strPropertyArray(i))
            'First deal with some properties that can not be obtained with Get.
            Case "name"
                strResult = objADS.Name
            Case "adspath"
                strResult = objADS.ADsPath
            Case Else
                strTemp = objADS.Get(strPropertyArray(i))
                If Err.Number Then
                    Err.Clear            'The property is not available.
                    If strPropertyArray(i) = "lastlogin" or strPropertyArray(i) = _
                        "lastlogoff" Then
                        strResult = CDate("1/1/1900")        'A date in the remote past.
                    Else
                        blnPrintProperties = True
                    End If
                Else
                    If IsArray(strResult) Then
                        Print strPropertyArray(i) & " is a multivalued property."
                        Print "The last value is used."
                        strResult = strResult(UBound(strTemp))
                    Else
                        strResult =  strTemp
                    End If
                End If
        End Select
        strOutput = strOutput & "         " & strResult
    Next
    WriteLine strOutput, objOutputFile

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

Sub Print(ByRef strMessage)
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
'********************************************************************
'    THE CODE BELOW IS FOR EVULUATING THE CRITERIA STRING ONLY
'********************************************************************
'********************************************************************

'********************************************************************
'*
'* Function intEvalCriteria()
'* Purpose: Evaluates a string that can contain brackets.
'* Input:   strString    a string
'* Output:  intEvalCriteria = CONST_FAILED if the evaluation failed,
'*          otherwise it is the result of the evaluation.
'* Example: If strString="(1 and 0) or 1" then intEvalCriteria=1.
'*
'********************************************************************

Private Function intEvalCriteria(ByVal strString)

    ON ERROR RESUME NEXT

    Dim intLeft, intRight, i, strTemp

    'Check the number of ")" and "("
    If intCharCount(strString, ")") <> intCharCount(strString, "(") Then
        intEvalCriteria = CONST_FAILED        'Incorrect syntax
        Exit Function
    End If

    'Now get rid of all double spaces.
    Do
        i = InStr(1, strString, "  ")
        If i Then
            'Replace double spaces with single ones.
            strString = Replace(strString, "  ", " ")
        End If
    Loop Until i = 0

    Do
        'Look for first ")" in the array
        intRight = InStr(1, strString, ")")
        If intRight = 0 Then
            'There is no quote in the array
            intEvalCriteria = intEvalNoQuote(strString)
            Exit Function
        End If


        intLeft = InStrRev(strString, "(", intRight, 1)
        If intLeft = 0 Then
            intEvalCriteria = CONST_FAILED        'Syntax error
            Exit Function
        End If

        strTemp = Mid(strString, intLeft+1, intRight-intLeft-1)

        If strTemp <> "" Then
            i = intEvalNoQuote(strTemp)
            If i = CONST_FAILED Then
                intEvalCriteria = i
                Exit Function
            End If
        Else
            i = ""
        End If

        If blnReplaceString(strString, intLeft, intRight, i) Then
            intEvalCriteria = CONST_FAILED
            Exit Function
        End If
    Loop Until Len(strString) = 1

    intEvalCriteria = CInt(strString)
    If Err.Number Then
        Err.Clear
        intEvalCriteria = CONST_FAILED
    End If

End Function

'********************************************************************
'*
'* Function blnReplaceString()
'* Purpose: Replaces a sub string in a string with another sub string.
'* Input:   strString       a string
'*          intStart        the starting position of the sub string
'*          intEnd          the ending position of the sub string
'*          strReplace      the new sub string
'* Output:  blnReplaceString = True if an error occurred and False otherwise.
'*
'********************************************************************

Private Function blnReplaceString(strString, intStart, intEnd, strReplace)

    Dim strLeft, strRight, intLen

    blnReplaceString = False    'No error.

    intLen = Len(strString)
    If intStart < 1 or intEnd > intLen Then
        blnReplaceString = True        'This is an error
        Exit Function
    End If

    strLeft = Left(strString, intStart-1)
    strRight = Right(strString, intLen-intEnd)
    strString = strLeft & strReplace & strRight

End Function

'********************************************************************
'*
'* Function intCharCount()
'* Purpose: Finds the number of times a character appears in a string.
'* Input:   strString   a string
'*          chr         a character
'* Output:  intCharCount is the number of times a character appears in an array.
'*
'********************************************************************

Private Function intCharCount(ByVal strString, ByVal chr)

    Dim i, strTemp

    i = Len(strString)
    strTemp = Replace(strString, chr, "")
    intCharCount = i - Len(strTemp)

End Function


'********************************************************************
'*
'* Function intEvalNoQuote()
'* Purpose: Evaluates a string that does not contain any quote.
'* Input:   strString    a string
'* Output:  intEvalNoQuote = CONST_FAILED if the evaluation failed,
'*          otherwise it is the result of the evaluation.
'* Example: If strString="1 and 0 or 1", then intEvalNoQuote=1.
'*
'********************************************************************

Private Function intEvalNoQuote(ByVal strString)

    ON ERROR RESUME NEXT

    Dim i, intLeft, intRight, blnLeft, blnRight, chrSpace
    Dim intAnd, intNot, intOr

    chrSpace = chr(32)
    strString = LCase(Trim(strString))

    'Handling all "Not"
    Do
        intNot = InStr(1, strString, "not")
        If intNot Then
            If intNot > (Len(strString)-4) Then
                intEvalNoQuote = CONST_FAILED    'It is an error.
                Exit Function
            End If
            intRight = InStr(intNot+4, strString, chrSpace)
            If intRight = 0 Then
                intRight = Len(strString) + 1
            End If
            blnRight = CBool(Mid(strString, intNot+4, intRight-intNot-4))
            'Commit the Not operation
            i = 1 + CInt(blnRight)        '1 for true and 0 for false
            If Err.Number Then
                Err.Clear
                intEvalNoQuote = CONST_FAILED    'An error occurred.
                Exit Function
            End If
            'Get the result into the string.
            If blnReplaceString(strString, intNot, intRight-1, i) Then
                intEvalNoQuote = CONST_FAILED    'An error occurred.
                Exit Function
            End If
        End If
    Loop Until intNot = 0

    'Handling all "and"
    Do
        intAnd = InStr(1, strString, "and")
        If intAnd Then
            If intAnd < 3 or intAnd > (Len(strString)-4) Then
                intEvalNoQuote = CONST_FAILED    'It is an error.
                Exit Function
            End If
            intLeft = InStrRev(strString, chrSpace, intAnd-2)
            intRight = InStr(intAnd+4, strString, chrSpace)
            If intLeft = 0 Then
                intLeft = 0
            End If
            If intRight = 0 Then
                intRight = Len(strString) + 1
            End If
            'Get the value to the left
            blnLeft = CBool(Mid(strString, intLeft+1, intAnd-intLeft-2))
            'Get the value to the right
            blnRight = CBool(Mid(strString, intAnd+4, intRight-intAnd-4))
            If Err.Number Then
                Err.Clear
                intEvalNoQuote = CONST_FAILED    'An error occurred.
                Exit Function
            End If
            'Commit the And operation
            i = -CInt(blnLeft and blnRight)            '1 for true and 0 for false
            If blnReplaceString(strString, intLeft+1, intRight-1, i) Then
                'Get the result into the string.
                intEvalNoQuote = CONST_FAILED    'An error occurred.
                Exit Function
            End If
        End If
    Loop Until intAnd = 0

    'Handling all "or"
    Do
        intOr = InStr(1, strString, "or")
        If intOr Then
            If intOr < 3 or intOr > (Len(strString)-3) Then
                intEvalNoQuote = CONST_FAILED    'It is an error.
                Exit Function
            End If
            intLeft = InStrRev(strString, chrSpace, intOr-2)
            intRight = InStr(intOr+3, strString, chrSpace)
            If intLeft = 0 Then
                intLeft = 0
            End If
            If intRight = 0 Then
                intRight = Len(strString) + 1
            End If
            'Get the value to the left
            blnLeft = CBool(Mid(strString, intLeft+1, intOr-intLeft-1))
            'Get the value to the right
            blnRight = CBool(Mid(strString, intOr+3, intRight-intOr-3))
            If Err.Number Then
                Err.Clear
                intEvalNoQuote = CONST_FAILED    'An error occurred.
                Exit Function
            End If
            'Commit the And operation
            i = -CInt(blnLeft or blnRight)            '1 for true and 0 for false
            If blnReplaceString(strString, intLeft+1, intRight-1, i) Then
                'Get the result into the string.
                intEvalNoQuote = CONST_FAILED    'An error occurred.
                Exit Function
            End If
        End If
    Loop Until intOr = 0

    strString = Trim(strString)
    If Len(strString) > 1 Then
        intEvalNoQuote = CONST_FAILED    'It is an error.
    Else
        intEvalNoQuote = CInt(strString)
    End If

End Function

'********************************************************************
'*
'* Note:
'*
'* 1. The criteria should be combinations of expressions like (property1:=value1).
'*
'* 2. In parsing the input the property names (eg, property1) are read into array
'*    strProperties, the values(eg, value1) are read into array strPropertyValues,
'*    and the operators(eg, =) are read into strOperators(0).
'*
'* 3. String strCriteria stores the criteria with each comparison property1:=value1
'*    replaced by an integer representing the order it appears in the criteria.
'*    For example, criteria "name:='j*' and lastlogin:>4/3/98 and lastlogin:<8/8/98"
'*    becomes "0 and 1 and 2". This expression is then evaluated to determine whether
'*    the syntax is correct.
'*
'* 4. For users who has never logged in,  the lastlogin and lastlogoff dates assigned
'*    are "1/1/1900".
'*
'*  File Name:    ChkUsers.vbs
'*
'*    A detailed description:
'*
'*    This script is intended to be used to check a domain or container for users
'*    satisfying a given criteria. But it can be easily adapted to check a domain or
'*    container for other types of objects satisfying a given criteria. The only change
'*    required for the script is to change the filter for the domain or container in
'*    function intChkOneDomain().
'*
'*    The input to the script includes the ADsPath of a domain or container. It is also
'*    possible to check multiple domains or containers if the criteria is the same for all
'*    of them. To do this simply save ADsPaths of these domains or containers into a text
'*    file, one in a line, and use the [/I:inputfile] option instead of [/A:adspath].
'*
'*    /C:criteria is the only mandatory input for this script. The criteria should be
'*    enclosed in double quotes to be interpreted correctly. The criteria is composed of
'*    many comparisons linked using logic operators, such as And, Or, Not. Each comparison
'*    is in the format of property:>value, where property is a valid property name of the
'*    user object and > can also be replaced by < or = and value is a valid value of the
'*    property. For example FullName:='John Smith' specifies the user's fullname to be
'*    "John Smith". Note that there is a colon between the property name and the
'*    comparison operator and a string with space should be enclosed in single quotes. If
'*    the comparison operator is omitted the default is =. It is also possible to use wild
'*    card * with operator =.
'*
'*    Brackets can be used in the criteria so the criteria experession will be in a form
'*    like ((A or B) and Not C) where each of A, B, C is a valid comparison in the form of
'*    property:>value. In function intParseCmdLine(), each comparison is replaced with an
'*    integer in the order it appears. For example the expression above becomes ((0 or 1)
'*    and Not 2). This expression is first evaluated to determine whether the syntax is
'*    correct. For example, a left bracket without a right bracket, or an incorrect usage
'*    of a comparison operator would trigger an "incorrect syntax" warning and the program
'*    execution would be terminated. The comparisons are stored in three string arrays:
'*    strProperties, strPropertyValues, strOperators. strProperties stores the names of
'*    properties, while strPropertyValues stores the corresponding values and strOperators
'*    stores the operators.
'*
'*    If the syntax is correct the script proceeds to list every user in the
'*    domain/container and for each user the comparisons are evaluted to be either 1(true)
'*    or 0(false). For example, FullName:=J* will be 1 if the user's fullname starts with
'*    letter J and 0 otherwise. The values of these comparisons are then plugged back into
'*    the criteria experission and the criteria is evaluated to be either true or false
'*    based on these values. For example, if A is true, B is False and C is False then the
'*    strCriteria becomes ((1 or 0) and Not 0) which evaluates to 1, or True.
'*
'*    If the criteria is evaluated to be true, properties as listed in [/P:properties] of
'*    the user will be retrieved and either saved into an output file or printed on the
'*    screen. If no property is specified using [/P:properties], ADsPath is used as the
'*    default.
'*
'********************************************************************

'********************************************************************
'*                                                                  *
'*                           End of File                            *
'*                                                                  *
'********************************************************************

'********************************************************************
'*
'* Procedures calling sequence: CHKUSERS.VBS
'*
'*  intChkProgram
'*  intParseCmdLine
'*      blnParseCriteria
'*      intEvalCriteria
'*  ShowUsage
'*  ConvertPropertyValues
'*  ChkUsers
'*      intChkOneDomain
'*          blnGetOneProperty
'*          intCompare
'*          blnCopyResults
'*          intEvalCriteria
'*          blnPrintProperties
'*              WriteLine
'*
'********************************************************************
