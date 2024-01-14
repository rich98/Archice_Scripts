
'********************************************************************
'*
'* File:        MODIFYUSERS.VBS
'* Created:     August 1998
'* Version:     1.0
'*
'* Main Function: Modifies attributes of one or more users.
'* Usage: MODIFYUSERS.VBS /A:adspath [/I:inputfile] [property1:propertyvalue1]
'*        [property2:propertyvalue2 ...] [/U:username] [/W:password] [/ALL] [/Q]
'*
'* Copyright (C) 1998 Microsoft Corporation
'*
'********************************************************************

OPTION EXPLICIT
ON ERROR RESUME NEXT

'Define constants
CONST CONST_ERROR                       = 0
CONST CONST_WSCRIPT                     = 1
CONST CONST_CSCRIPT                     = 2
CONST CONST_SHOW_USAGE                  = 3
CONST CONST_PROCEED                     = 4
CONST CONST_STRING_NOT_FOUND            = -1
CONST CONST_UF_PASSWORD_CANT_CHANGE     = 64            'constants for setting user flags
CONST CONST_UF_PASSWORD_CAN_CHANGE      = 131007
CONST CONST_UF_DONT_EXPIRE_PASSWORD     = 65536
CONST CONST_UF_DO_EXPIRE_PASSWORD       = 65535

'Declare variables
Dim strDomain, strFile, strCurrentUser, strPassword, blnQuiet, intOpMode, blnAll, i
ReDim strArgumentArray(0), strPropertyArray(0), strPropertyValueArray(0)

'Initialize variables
intOpMode = 0
blnQuiet = False
blnAll = False                'By default do not change attributes of all users
strDomain = ""
strFile = ""
strCurrentUser = ""
strPassword = ""
strArgumentArray(0) = ""
strPropertyArray(0) = ""
strPropertyValueArray(0) = ""

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
            "1. Using ""CScript MODIFYUSERS.vbs arguments"" for Windows 95/98 or" & vbCRLF & _
            "2. Changing the default Windows Scripting Host setting to CScript" & vbCRLF & _
            "    using ""CScript //H:CScript //S"" and running the script using" & vbCRLF & _
            "    ""MODIFYUSERS.vbs arguments"" for Windows NT."
        WScript.Quit
    Case Else
        WScript.Quit
End Select

'Parse the command line
intOpMode = intParseCmdLine(strArgumentArray, strDomain, strFile, blnAll, strCurrentUser,_
        strPassword, blnQuiet, strPropertyArray, strPropertyValueArray)
If Err.Number Then
    Print "Error 0X" & CStr(Hex(Err.Number)) & " occurred in parsing the command line."
    If Err.Description <> "" Then
        Print "Error description: " & Err.Description & "."
    End If
    WScript.quit
End If

Select Case intOpMode
    Case CONST_SHOW_USAGE
        Call ShowUsage()
    Case CONST_PROCEED
        Call ModifyUsers(strDomain, strFile, blnAll, strCurrentUser,_
             strPassword, blnQuiet, strPropertyArray, strPropertyValueArray)
    Case CONST_ERROR
        'Do nothing.
    Case Else
        Wscript.Echo "Error occurred in passing parameters."
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
'* Output:  strDomain           the ADsPath of the domain
'*          strFile             the input file name including the path
'*          blnAll              specifies whether the operation is over the whole domain
'*          strCurrentUser      the name or cn of the current user
'*          strPassword         the current user password
'*          blnQuiet            specifies whether to suppress messages
'*          strPropertyArray    an array of user properties names
'*          strPropertyValueArray    an array of the corresponding user properties
'*          intParseCmdLine     is set to one of CONST_ERROR, CONST_SHOW_USAGE, CONST_PROCEED.
'*
'********************************************************************

Private Function intParseCmdLine(strArgumentArray, strDomain, strFile, blnAll, strCurrentUser,_
        strPassword, blnQuiet, strPropertyArray, strPropertyValueArray)

    ON ERROR RESUME NEXT

    Dim i, j, strFlag

    strFlag = strArgumentArray(0)
    If strFlag = "" then                    'No arguments have been received
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

    strDomain = strFlag     'The first parameter must be the domain ADsPath.

    j = 0
    For i = 1 to UBound(strArgumentArray)
        strFlag = Left(strArgumentArray(i), InStr(1, strArgumentArray(i), ":")-1)
        If Err.Number Then            'An error occurs if there is no : in the string
            Err.Clear
            If LCase(strArgumentArray(i)) = "/all" Then
                blnAll = True
            ElseIf LCase(strArgumentArray(i)) = "/q" Then
                blnQuiet = True
            Else
                Print strArgumentArray(i) & " is not recognized as a valid input."
                intParseCmdLine = CONST_ERROR
                Exit Function
            End If
        Else
            Select Case LCase(strFlag)
                Case "/i"
                    strFile = Right(strArgumentArray(i), Len(strArgumentArray(i))-3)
                Case "/u"
                    strCurrentUser = Right(strArgumentArray(i), Len(strArgumentArray(i))-3)
                Case "/w"
                    strPassword = Right(strArgumentArray(i), Len(strArgumentArray(i))-3)
                Case else
                    ReDim Preserve strPropertyArray(j), strPropertyValueArray(j)
                    strPropertyArray(j) = strFlag
                    strPropertyValueArray(j) = Right(strArgumentArray(i), _
                        Len(strArgumentArray(i))-InStr(1, strArgumentArray(i), ":"))
                    If strPropertyValueArray(j) = ""  Then
                        Print "Warning: property " & strFlag & " does not have a value!"
                    End If
                    j = j + 1
            End Select
        End If
    Next

    If (strDomain = "") Then                                        'strDomain is required
        Print "The ADsPath of the domain is missing."
        intParseCmdLine = CONST_ERROR
        Exit Function
    End If

    If blnAll Then
        If strFile <> "" Then
            Wscrip.Echo "The attributes of all users in the domain will be modified. File " _
                & strFile & " is ignored."
        End If
    ElseIf (strFile = "") and (strPropertyArray(0) = "" = "") Then
        Print "The user account name is missing."
        intParseCmdLine = CONST_ERROR
        Exit Function
    End If

    intParseCmdLine = CONST_PROCEED

End Function

'********************************************************************
'*
'* Sub ShowUsage()
'* Purpose:   Shows the correct usage to the user.
'* Input:     None
'* Output:    Help messages are displayed on screen.
'*
'********************************************************************

Sub ShowUsage()

    Wscript.Echo ""
    Wscript.Echo "Modifies attributes of one or more users."  & vbCRLF
    Wscript.Echo "MODIFYUSERS.VBS adspath [/I:inputfile] [property1:propertyvalue1]"
    Wscript.Echo "[property2:propertyvalue2...] [/U:username] [/W:password] [/ALL] [/Q]"
    Wscript.echo "   /I, /U, /W"
    Wscript.Echo "                 Parameter specifiers."
    Wscript.echo "   adspath       ADsPath of a user object container."
    Wscript.echo "   inputfile     A text file with each line in the following format:"
    Wscript.echo "                 property1:propertyvalue1 property2:propertyvalue2 ..."
    Wscript.echo "   property[i], propertyvalue[i]"
    Wscript.echo "                 Name and value of a user property."
    Wscript.echo "   username      Username of the current user."
    Wscript.echo "   password      Password of the current user."
    Wscript.echo "   /ALL          Resets attributes of all users in a domain."
    Wscript.echo "   /Q            Suppresses all output messages." & vbCRLF
    Wscript.Echo "EXAMPLE:"
    Wscript.echo "MODIFYUSERS.VBS WinNT://FooFoo name:jsmith"
    Wscript.echo "   passwordexpired:1 description:""FooFoo domain users"""
    Wscript.echo "   sets user jsmith's password to expired and changes the"
    Wscript.echo "   description of the user to ""FooFoo domain users""."

End Sub

'********************************************************************
'*
'* Sub ModifyUsers()
'* Purpose: Modifies attributes of multiple users.
'* Input:   strDomain           the ADsPath of the domain
'*          strFile             the input file name including the path
'*          blnAll              specifies whether the operation is over the whole domain
'*          strCurrentUser      the name or cn of the current user
'*          strPassword         the current user password
'*          blnQuiet            specifies whether to suppress messages
'*          strPropertyArray    an array of user properties names
'*          strPropertyValueArray    an array of the corresponding user properties
'* Output:  None
'*
'********************************************************************

Sub ModifyUsers(strDomain, strFile, blnAll, strCurrentUser,_
    strPassword, blnQuiet, strPropertyArray, strPropertyValueArray)

    ON ERROR RESUME NEXT

    Dim strProvider, objDomain, strUser, objUser, i, objFileSystem, objInputFile, strInput
    Dim blnResult

    'Check the provider passed
    strProvider = Left(strDomain, InStr(1, strDomain, ":")-1)
    If Err.Number Then
        Print "The ADsPath " & strDomain & " of the container object is incorrect!"
        Err.Clear
        Exit Sub
    End If

    If strProvider <> "WinNT" and strProvider <> "LDAP" Then
        Print "The provider " & strProvider & " is not supported."
        Exit Sub
    End If

    Print "Getting domain " & strDomain & "..."
    If strCurrentUser = "" Then            'no user credential is passed
        Set objDomain = GetObject(strDomain)
    Else
        Set objProvider = GetObject(strProvider & ":")
		'Use user authentication
        Set objDomain = objProvider.OpenDsObject(strDomain,strCurrentUser,strPassword,1)
    End If
    If Err.Number then
		If CStr(Hex(Err.Number)) = "80070035" Then
			Print "Object " & strDomain & " is not found."
		Else
			Print "Error 0x" & CStr(Hex(Err.Number)) & " occurred in getting object " _
				& strDomain & "."
			If Err.Description <> "" Then
				Print "Error description: " & Err.Description & "."
			End If
		End If
		Err.Clear
        Exit Sub
    End If

    If blnAll Then        'we need to change attributes of each user in the domain
        Print "Modifying attributes of all users in domain " & strDomain & "."
        'Make sure that attributes such as name, cn, samaccountname are not in
        'the attribute list
        If strProvider = "WinNT" Then
            'This deletes the user's name from the list so they can not be modified.
            Call strGetUser("name", strPropertyArray, strPropertyValueArray)
        Else                'must be LDAP
            'This deletes the user's samaccountname and cn from the list so
            'they can not be modified.
            Call strGetUser("samaccountname", strPropertyArray, strPropertyValueArray)
            Call strGetUser("cn", strPropertyArray, strPropertyValueArray)
        End If

        objDomain.Filter = Array("user")
        For Each objUser in objDomain
            If strProvider = "WinNT" Then
                strUser = objUser.Name
            Else                'must be LDAP
                strUser = "CN=" & objUser.CN
            End If
            Call ModifyOneUser(objDomain, strProvider, strUser, strPropertyArray, _
                 strPropertyValueArray)
        Next
        Exit Sub
    End If

    If strPropertyArray(0) <> "" Then        'need to modify attributes for one user
        If strProvider = "WinNT" Then
            'This deletes the user's name from the list so they can not be modified.
            strUser = strGetUser("name", strPropertyArray, strPropertyValueArray)
        Else                'must be LDAP
            'This deletes the user's samaccountname and cn from the list
            'so they can not be modified.
            Call strGetUser("samaccountname", strPropertyArray, strPropertyValueArray)
            strUser = "CN=" & strGetUser("cn", strPropertyArray, strPropertyValueArray)
        End If
        Call ModifyOneUser(objDomain, strProvider, strUser, strPropertyArray, _
             strPropertyValueArray)
    End If

    If strFile <> "" Then
        'Create a filesystem object
        set objFileSystem = CreateObject("Scripting.FileSystemObject")
        If Err.Number Then
            Print "Error 0X" & CStr(Hex(Err.Number)) & _
                " occurred in creating a filesystem object."
            If Err.Description <> "" Then
                Print "Error description: " & Err.Description & "."
            End If
            Exit Sub
        End If

        'Opens a file for input
        set objInputFile = objFileSystem.OpenTextFile(strFile)
        If Err.Number Then
            Print "Error 0X" & CStr(Hex(Err.Number)) & " occurred in opening file " & strFile
            If Err.Description <> "" Then
                Print "Error description: " & Err.Description & "."
            End If
            Exit Sub
        End If

        'Read from the file
        i = 0
        While not objInputFile.AtEndOfStream
            strInput = Trim(objInputFile.ReadLine)    'Get rid of leading and trailing spaces
            If Not (strInput = "") Then
                blnResult = blnParseInputFile(strInput, strPropertyArray, _
                            strPropertyValueArray)
            End If
            If blnResult Then
                Print "Error occurred in parsing the input line " & vbCRLF & strUser & "."
            Else
                If strPropertyArray(0) <> "" Then
                    If strProvider = "WinNT" Then
                        strUser = strGetUser("name", strPropertyArray, _
                                  strPropertyValueArray)
                    Else                'must be LDAP
                        'The samaccountname is not to be modified
                        Call strGetUser("samaccountname", strPropertyArray, _
                             strPropertyValueArray)
                        strUser = "CN=" & strGetUser("cn", strPropertyArray, _
                                strPropertyValueArray)
                    End If

                    If strUser = "" Then
                        Print "Warning: The user name is not found in the input file."
                        Exit Sub
                    Else
                        Call ModifyOneUser(objDomain, strProvider, strUser, _
                             strPropertyArray, strPropertyValueArray)
                    End If
                End If
            End If
        Wend
        objInputFile.Close
    End If

End Sub

'********************************************************************
'*
'* Function blnParseInputFile()
'* Purpose: Parses a line of input from the input file.
'* Input:   strInput                a string to be parsed
'* Output:  strPropertyArray        an array of user properties names
'*          strPropertyValueArray   an array of the corresponding user properties
'*          blnParseInputFile is set to True if an error occurred and False otherwise.
'*
'********************************************************************

Function blnParseInputFile(strInput, strPropertyArray, strPropertyValueArray)

    ON ERROR RESUME NEXT

    Dim strSpace, strQuote, strColon, i, intSpace, intQuote, intColon

    strSpace = chr(32)                'space
    strQuote = chr(34)                'double quote
    strColon = chr(58)                'colon
    blnParseInputFile = False         'No error

    i = 0
    Do While Len(strInput)        'if strInput is not empty
        ReDim Preserve strPropertyArray(i), strPropertyValueArray(i)
        'The property name is up to the first colon
        intColon = InStr(1, strInput, strColon)
        If intColon = 0 Then    'There is no colon in the input line.
            blnParseInputFile = True        'This is an error
            Exit Do
        End If
        strPropertyArray(i) = Trim(Left(strInput, intColon-1))
        strInput = Trim(Right(strInput, Len(strInput)-intColon))
        If InStr(1, strPropertyArray(i), strQuote) or _
            InStr(1, strPropertyArray(i), strSpace)    or _
            InStr(1, strPropertyArray(i), strColon) or _
            strInput = "" or strPropertyArray(i) = "" Then
            blnParseInputFile = True        'This is an error.
            Exit Do
        End If

        'If there is a quote for this property value
        If Left(strInput, 1) = strQuote Then
            'The property value is from the first quote to the second quote
            intQuote = InStr(2, strInput, strQuote)
            If intQuote = 0 Then        'There is no second quote in the string.
                blnParseInputFile = True        'This is an error
                Exit Do
            End If
            strPropertyValueArray(i) = Trim(Mid(strInput, 2, intQuote-2))
            strInput = Trim(Right(strInput, Len(strInput)-intQuote))
        Else
            'If this property value does not start with a quote it must end with a space
            'unless it is at the end of the input string
            intSpace = InStr(1, strInput, strSpace)
            If intSpace = 0 Then        'There is no space in the string.
                'Simply assign strInput to the property value.
                strPropertyValueArray(i) = strInput
                strInput = ""            'The allows the loop to come to a stop normally.
            Else
                'The property value is up to the first space
                strPropertyValueArray(i) = Left(strInput, intSpace-1)
                strInput = Trim(Right(strInput, Len(strInput)-intSpace))
            End If
        End If
        i = i + 1
    Loop

End Function

'********************************************************************
'*
'* Function strGetUser()
'* Purpose: Searches for an element in strArray1 and strArray2.
'* Input:   strArray1       an array of user properties names
'*          strArray2       an array of the corresponding user properties
'* Output:  If strTarget    is found in strArray1 as element i then strGetUser is set to
'*          strArray2(i)    and then the i-th element of both strArray1 and
'*                          strArray2 are deleted.
'*          Otherwise strGetUser = "" and strArray1 and strArray2 are unchanged.
'*
'********************************************************************

Private Function strGetUser(ByVal strTarget, strArray1, strArray2)

    Dim i

    i = intSearchArray(strTarget, strArray1)
    If i = CONST_STRING_NOT_FOUND Then
        strGetUser = ""
    Else
        strGetUser = strArray2(i)
        Call DeleteOneElement(i, strArray1)
        Call DeleteOneElement(i, strArray2)
    End If

End Function

'********************************************************************
'*
'* Sub ModifyOneUser()
'* Purpose: Modifies attributes of one user.
'* Input:   objDomain               a domain object
'*          strProvider             an ADSI provider name
'*          strUser                 the username or cn of the user to be deleted
'*          strPropertyArray        an array of user properties names.
'*          strPropertyValueArray   an array of the corresponding user properties
'* Output:  None
'*
'********************************************************************

Sub ModifyOneUser(objDomain, strProvider, strUser, strPropertyArray, strPropertyValueArray)

    ON ERROR RESUME NEXT

    Dim objUser, lngFlag, i, j

    strUser = LCase(strUser)        'make sure that the user name is lower cased

'    Check whether the user already exists
    set objUser = objDomain.GetObject("user", strUser)
    If Err.Number Then        'The user does not exist.
        Print "User " & strUser & " does not exist in domain " & strDomain & "."
        Err.Clear
        Exit Sub
    End If

    If Not (strPropertyArray(0) = "") Then
        Print Time & " modifying attributes of user " & strUser
        If strProvider = "WinNT" Then
            lngFlag = objUser.Get("UserFlags")
        Else                'must be LDAP
            lngFlag = objUser.Get("UserAccountControl")
        End If
        If Err.Number Then
            Print "Error 0X" & CStr(Hex(Err.Number)) & _
                  " occurred in getting userflags for user " & strUser & "."
            Err.Clear
            Exit Sub
        End If

        For i = 0 To UBound(strPropertyArray)
            'First let's deal with several special properties
            Select Case LCase(strPropertyArray(i))
                Case "password"
                    objUser.SetPassword CStr(strPropertyValueArray(i))
                Case "passwordexpired"
                    If CBool(strPropertyValueArray(i)) Then
                        'First we need to make sure that no conflict exists
                        'now user can change password
                        lngFlag = lngFlag and CONST_UF_PASSWORD_CAN_CHANGE
                        'this sets the password to expire
                        lngFlag = lngFlag and CONST_UF_DO_EXPIRE_PASSWORD
                        If strProvider = "WinNT" Then
                            objUser.put "userFlags", CLng(lngFlag)
                            objUser.Put "PasswordExpired", CLng(1)
                        Else                'must be LDAP
                            objUser.put "UserAccountControl", CLng(lngFlag)
                            objUser.put "pwdLastSet", CLng(0)
                        End If
                        Print "        " & "User can change password!"
                        Print "        " & "Password can expire!"
                        Print "        PasswordExpired = True"
                    Else
                        If strProvider = "WinNT" Then
                            objUser.Put "PasswordExpired", CLng(0)
                        Else                'must be LDAP
                            objUser.put "pwdLastSet", CLng(-1)
                        End If
                        Print "        PasswordExpired = False"
                    End If
                Case "accountdisabled"
                    If CBool(strPropertyValueArray(i)) Then
                        objUser.AccountDisabled = True
                        Print "        AccountDisabled = True"
                    Else
                        objUser.AccountDisabled = False
                        Print "        AccountDisabled = False"
                    End If
                Case "accountexpirationdate"
                    If IsDate(strPropertyValueArray(i)) Then
                        If DateDiff("d", now, CDate(strPropertyValueArray(i))) < 2 Then
                            Print "        Expiration date is too close."
                        Else
                            objUser.AccountExpirationDate = CDate(strPropertyValueArray(i))
                            Print "        AccountExpirationDate = " & _
                                  CDate(strPropertyValueArray(i))
                        End If
                    Else
                        Print "        Warning: " & strPropertyValueArray(i) & _
                              " is not a valid date."
                        Print "        The expiration date is not set."
                    End If
                Case "accountlockout"
                    If CBool(strPropertyValueArray(i)) Then
                        Print "        You can not set the user's account lockout to be true."
                    Else
                        'This is the default so nothing needs to be done
                        'objUser.IsAccountLocked = False
                    End If
                Case "usercannotchangepassword"
                    If CBool(strPropertyValueArray(i)) Then
                        If strProvider = "WinNT" Then
                            'Make sure there is no conflict
                            objUser.Put "PasswordExpired", CLng(0)
                        Else                'must be LDAP
                            objUser.put "pwdLastSet", CLng(-1)
                        End If
                        'now user can not change password
                        lngFlag = lngFlag or CONST_UF_PASSWORD_CANT_CHANGE
                        Print "        PasswordExpired = False"
                        Print "        " & "User can not change password!"
                    Else
                        'now user can change password
                        lngFlag = lngFlag and CONST_UF_PASSWORD_CAN_CHANGE
                        Print "        " & "User can change password!"
                    End If
                    If strProvider = "WinNT" Then
                        objUser.put "userFlags", CLng(lngFlag)
                    Else                'must be LDAP
                        objUser.put "UserAccountControl", CLng(lngFlag)
                    End If
                Case "passwordneverexpires"
                    If strPropertyValueArray(i) Then
                        If strProvider = "WinNT" Then
                            'Make sure there is no conflict
                            objUser.Put "PasswordExpired", CLng(0)
                        Else                'must be LDAP
                            objUser.put "pwdLastSet", CLng(-1)
                        End If
                        'this sets the password to never expires
                        lngFlag = lngFlag or CONST_UF_DONT_EXPIRE_PASSWORD
                        Print "        PasswordExpired = False"
                        Print "        " & "Password never expires!"
                    Else
                        'this sets the password to expire
                        lngFlag = lngFlag and CONST_UF_DO_EXPIRE_PASSWORD
                        Print "        " & "Password can expire!"
                    End If
                    If strProvider = "WinNT" Then
                        objUser.put "userFlags", CLng(lngFlag)
                    Else                'must be LDAP
                        objUser.put "UserAccountControl", CLng(lngFlag)
                    End If
                Case Else
                    Print "        " & strPropertyArray(i) & " = " &  _
                           CStr(strPropertyValueArray(i))
                    objUser.Put strPropertyArray(i), CStr(strPropertyValueArray(i))
            End Select
            If Err.Number Then
                Print "Error 0X" & CStr(Hex(Err.Number)) & " occurred in modifying property "_
                      & strPropertyArray(i) & " for user " & strUser & "."
                If Err.Description <> "" Then
                    Print "Error description: " & Err.Description & "."
                End If
                Err.Clear
            End If
        Next
    End If

    objUser.SetInfo                'commit the changes
    If Err.Number Then
        Print "Error 0X" & CStr(Hex(Err.Number)) & _
              " occurred in modifying attributes of user " & strUser & "."
        If Err.Description <> "" Then
            Print "Error description: " & Err.Description & "."
        End If
        Err.Clear
    Else
        Wscript.Echo "Succeeded in modifying attributes of user " & strUser & "."
    End If

End Sub

'********************************************************************
'*
'* Sub DeleteOneElement()
'* Purpose: Deletes one element from an array.
'* Input:   i           the index of the element to be deleted
'*          strArray    the array to work on
'* Output:  strArray    the array with the i-th element deleted
'*
'********************************************************************

Private Sub DeleteOneElement(ByVal i, strArray)

    Dim j, intUbound

    If Not IsArray(strArray) Then
        Print "Argument is not an array!"
        Exit Sub
    End If

    intUbound = UBound(strArray)

    If i > intUBound or i < 0 Then
        Print "Array index out of range!"
        Exit Sub
    ElseIf i < intUBound Then
        For j = i To intUBound - 1
            strArray(j) = strArray(j+1)
        Next
        j = j - 1
    Else                            'i = intUBound
        If intUBound = 0 Then        'There is only one element in the array
            strArray(0) = ""        'set it to empty
            j = 0
        Else                        'Need to delete the last element (i-th element)
            j = intUBound - 1
        End If
    End If

    ReDim Preserve strArray(j)

End Sub

'********************************************************************
'*
'* Function intSearchArray()
'* Purpose: Searches an array for a given string.
'* Input:   strTarget    the string to look for
'*          strArray    an array of strings to search against
'* Output:  If a match is found intSearchArray is set to the index of the element,
'*          otherwise it is set to CONST_STRING_NOT_FOUND.
'*
'********************************************************************

Private Function intSearchArray(ByVal strTarget, ByVal strArray)

    Dim i

    intSearchArray = CONST_STRING_NOT_FOUND

    If Not IsArray(strArray) Then
        Print "Argument is not an array!"
        Exit Function
    End If

    strTarget = LCase(strTarget)
    For i = 0 To UBound(strArray)
        If LCase(strArray(i)) = strTarget Then
            intSearchArray = i
        End If
    Next

End Function

'********************************************************************
'*
'* Sub Print()
'* Purpose:   Prints a message on screen if blnQuiet = False.
'* Input:     strMessage    the string to print
'* Output:    strMessage is printed on screen if blnQuiet = False.
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
'* Procedures calling sequence: MODIFYUSERS.VBS
'*
'*  intChkProgram
'*	intParseCmdLine
'*	ShowUsage
'*	ModifyUsers
'*      blnParseInputFile
'*		strGetUser
'*			intSearchArray
'*			DeleteOneElement
'*		ModifyOneUser
'*
'********************************************************************
