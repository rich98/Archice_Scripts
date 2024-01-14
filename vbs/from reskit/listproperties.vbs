
'********************************************************************
'*
'* File:        LISTPROPERTIES.VBS
'* Created:     August 1998
'* Version:     1.0
'*
'* Main Function: Lists properties of a given ADS object
'* Usage: LISTPROPERTIES.VBS adspath [/O:outputfile] [/U:username] [/W:password] [/Q]
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
Dim blnQuiet, i, intOpMode
ReDim strArgumentArray(0)

'Initialize variables
strArgumentArray(0) = ""
blnQuiet = False
strADsPath = ""
strUserName = ""
strPassword = ""
strOutputFile = ""

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
            "1. Using ""CScript LISTPROPERTIES.vbs arguments"" for Windows 95/98 or" _
                & vbCRLF & _
            "2. Changing the default Windows Scripting Host setting to CScript" & vbCRLF & _
            "    using ""CScript //H:CScript //S"" and running the script using" & vbCRLF & _
            "    ""LISTPROPERTIES.vbs arguments"" for Windows NT."
        WScript.Quit
    Case Else
        WScript.Quit
End Select

'Parse the command line
intOpMode = intParseCmdLine(strArgumentArray, strADsPath, _
            blnQuiet, strUserName, strPassword, strOutputFile)
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
        Call ListProperties(strADsPath, strUserName, strPassword, strOutputFile)
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
'* Output:  strADsPath          ADsPath of an ADs object
'*          strUserName         name of the current user
'*          strPassword         password of the current user
'*          strOutputFile       an output file name
'*          blnQuiet            specifies whether to suppress messages
'*          intParseCmdLine     is set to one of CONST_ERROR, CONST_SHOW_USAGE, CONST_PROCEED.
'*
'********************************************************************

Private Function intParseCmdLine(strArgumentArray, strADsPath, _
        blnQuiet, strUserName, strPassword, strOutputFile)

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

    strADsPath = strFlag    'The first parameter must be the ADsPath.

    For i = 1 to UBound(strArgumentArray)
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
                Case "/u"
                    strUserName = Right(strArgumentArray(i), Len(strArgumentArray(i))-3)
                Case "/o"
                    strOutputFile = Right(strArgumentArray(i), Len(strArgumentArray(i))-3)
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

    Wscript.Echo ""
    Wscript.Echo "Lists properties of a given ADS object." & vbCRLF
    Wscript.Echo "LISTPROPERTIES.VBS adspath  [/O:outputfile]"
    Wscript.Echo "[/U:username] [/W:password] [/Q]"
    Wscript.Echo "   /O, /U, /W    Parameter specifiers."
    Wscript.Echo "   adspath       The ADsPath of an ADs object."
    Wscript.Echo "   outputfile    The output file name."
    Wscript.Echo "   username      Username of the current user."
    Wscript.Echo "   password      Password of the current user."
    Wscript.Echo "   /Q            Suppresses all output messages." & vbCRLF
    Wscript.Echo "EXAMPLE:"
    Wscript.Echo "LISTPROPERTIES.VBS ""WinNT://FooFoo"""
    Wscript.Echo "   lists all properties of FooFoo object "

End Sub

'********************************************************************
'*
'* Sub ListProperties()
'* Purpose: Lists properties of a given ADS object.
'* Input:   strADsPath        ADsPath of an ADs object
'*          strUserName       name of the current user
'*          strPassword       password of the current user
'*          strOutputFile     an output file name
'* Output:  Properties of the object are either printed on screen or saved
'*          in strOutputFile.
'*
'********************************************************************

Private Sub ListProperties(strADsPath, strUserName, strPassword, strOutputFile)

    ON ERROR RESUME NEXT

    Dim strProvider, objProvider, objADs, objFileSystem, objOutputFile
    Dim objSchema, strProperty, strClassArray, strMessage, i
    ReDim strClassArray(0)

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
        Print "Could not find schema for object """ & strADsPath & """."
        Print "No property is found."
        Exit Sub
    End If

    'Open a file to output if strOutputFile is not empty.
    If strOutputFile = "" Then
        objOutputFile = ""
    Else
        'After discovering the object, open a file to save the results
        'Create a file object.
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

    strMessage = objADs.Schema
    If Err.Number Then
        Err.Clear
        ElseIf strMessage <> "" Then
        WriteLine "Schema = " &  strMessage, objOutputFile
    End If

    strMessage = objSchema.Name
    If Err.Number Then
        Err.Clear
        ElseIf strMessage <> "" Then
        WriteLine "Class = " &  strMessage, objOutputFile
    End If

    strMessage = objSchema.CLSID
    If Err.Number Then
        Err.Clear
        ElseIf strMessage <> "" Then
        WriteLine "CLSID = " &  strMessage, objOutputFile
    End If

    strMessage = objSchema.OID
    If Err.Number Then
        Err.Clear
        ElseIf strMessage <> "" Then
        WriteLine "OID = " &  strMessage, objOutputFile
    End If

    strMessage = objSchema.Abstract
    If Err.Number Then
        Err.Clear
        ElseIf strMessage <> "" Then
        WriteLine "Abstract = " &  strMessage, objOutputFile
    End If

    strMessage = objSchema.Auxilliary
    If Err.Number Then
        Err.Clear
        ElseIf strMessage <> "" Then
        WriteLine "Auxilliary = " &  strMessage, objOutputFile
    End If

    strMessage = objSchema.Container
    If Err.Number Then
        Err.Clear
        ElseIf strMessage <> "" Then
        WriteLine "Container = " &  strMessage, objOutputFile
    End If

    strMessage = objSchema.HelpFileName
    If Err.Number Then
        Err.Clear
        ElseIf strMessage <> "" Then
        WriteLine "HelpFileName = " &  strMessage, objOutputFile
    End If

    strMessage = objSchema.HelpFileContext
    If Err.Number Then
        Err.Clear
        ElseIf strMessage <> "" Then
        WriteLine "HelpFileContext = " &  strMessage, objOutputFile
    End If

    strMessage = objSchema.PrimaryInterface
    If Err.Number Then
        Err.Clear
        ElseIf strMessage <> "" Then
        WriteLine "PrimaryInterface = " &  strMessage, objOutputFile
    End If

    'Get all mandatory properties
    for each strProperty in objSchema.MandatoryProperties
        Call  GetOneProperty(objADs, strProperty, objOutputFile)
    next
    'Get all optional properties
    for each strProperty in objSchema.OptionalProperties
        Call  GetOneProperty(objADs, strProperty, objOutputFile)
    next
    for each strProperty in objSchema.NamingProperties
        Call  GetOneProperty(objADs, strProperty, objOutputFile)
    next
    for each strProperty in objSchema.DerivedFrom
        Call  GetOneProperty(objADs, strProperty, objOutputFile)
    next
    for each strProperty in objSchema.AuxDerivedFrom
        Call  GetOneProperty(objADs, strProperty, objOutputFile)
    next
    for each strProperty in objSchema.PossibleSuperiors
        Call  GetOneProperty(objADs, strProperty, objOutputFile)
    next
    for each strProperty in objSchema.Containment
        WriteLine "Possible containment = " &  strProperty
    next
    for each strProperty in objSchema.Qualifiers
        Call  GetOneProperty(objADs, strProperty, objOutputFile)
    next

    If strOutputFile <> "" Then
        Wscript.Echo "Results are saved in file " & strOutputFile & "."
        objOutputFile.Close
    End If

End Sub

'********************************************************************
'*
'* Sub GetOneProperty()
'* Purpose: Lists one property of a given ADS object.
'* Input:   objADS          an ADS object
'*          strProperty     name of a property
'*          objOutputFile   an output file object
'* Output:  The values the object property are either printed on screen or saved
'*          in objOutputFile.
'*
'********************************************************************

Sub GetOneProperty(objADS, strProperty, objOutputFile)

    ON ERROR RESUME NEXT

    Dim strResult, i, intUBound

    intUBound = 0

    strResult = objADS.Get(strProperty)

    If Err.Number Then
        Err.Clear            'The property is not available.
    Else
        If IsArray(strResult) Then
            intUBound = UBound(strResult)
            If (intUBound > 0) Then
                For i = 0 to UBound(strResult)
                    WriteLine (i+1) & " " & strProperty &  " = " & strResult(i), objOutputFile
                Next
            ElseIf strResult(0) <> "" Then
                WriteLine strProperty &  " = " & strResult(0), objOutputFile
            End If
        Else
            WriteLine strProperty &  " = " & strResult, objOutputFile
        End If
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
'* Procedures calling sequence: LISTPROPERTIES.VBS
'*
'*  intChkProgram
'*  intParseCmdLine
'*  ShowUsage
'*  ListProperties
'*      GetOneProperty
'*          WriteLine
'*
'********************************************************************
