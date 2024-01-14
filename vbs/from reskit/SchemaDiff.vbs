
'********************************************************************
'*
'* File:        SCHEMADIFF.VBS
'* Created:     April 1999
'* Version:     1.0
'*
'* Main Function: Compares the schema between two different enterprises
'* Usage: SCHEMADIFF.VBS [/U:username] [/W:password] [/D:domain] [/S:server] [/F:file name] [/Q]
'*
'* Copyright (C) 1999 Microsoft Corporation
'*
'********************************************************************

OPTION EXPLICIT
ON ERROR RESUME NEXT

'Define constants
CONST CONST_STRING_NOT_FOUND            = -1
CONST CONST_ERROR                       = 0
CONST CONST_WSCRIPT                     = 1
CONST CONST_CSCRIPT                     = 2
CONST CONST_SHOW_USAGE                  = 3
CONST CONST_PROCEED                     = 4

CONST ForWriting			= 2
CONST ForAppending			= 8

CONST ADS_OBJECT_NOTFOUND               = &H80072030
CONST ADS_ATTRIBUTE_NOTFOUND		= &H8007000A

'Declare variables
Dim strFile, strCurrentUser, strPassword, blnQuiet, intOpMode, i
Dim strArgumentArray, strServer
ReDim strArgumentArray(0)

'Initialize variables
intOpMode = 0
blnQuiet = False
strFile = ""
strCurrentUser = ""
strPassword = ""
strServer = ""
strArgumentArray(0) = ""


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
            "1. Using ""CScript DISPLAYOLD.VBS arguments"" for Windows 95/98 or" & vbCRLF & _
            "2. Changing the default Windows Scripting Host setting to CScript" & vbCRLF & _
            "    using ""CScript //H:CScript //S"" and running the script using" & vbCRLF & _
            "    ""SCHEMADIFF.VBS arguments"" for Windows NT."
        WScript.Quit
    Case Else
        WScript.Quit
End Select

'Parse the command line
intOpMode = intParseCmdLine(strArgumentArray, strCurrentUser,strPassword, blnQuiet, strServer, strFile)
If Err.Number Then
    WScript.Echo "Error 0x" & CStr(Hex(Err.Number)) & " occurred in parsing the command line."
    If Err.Description <> "" Then  WScript.Echo "Error description: " & Err.Description 
    WScript.quit
End If

Select Case intOpMode
    Case CONST_SHOW_USAGE
        Call ShowUsage()
    Case CONST_PROCEED
        Call SCHEMADIFF(strServer, strCurrentUser, strPassword, blnQuiet, strFile)
    Case CONST_ERROR
        'Do nothing.
    Case Else
        Wscript.Echo "Error occurred in passing parameters."
End Select

WScript.Quit

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
        WScript.Echo "Error 0x" & CStr(Hex(Err.Number)) & " occurred."
        If Err.Description <> "" Then WScript.Echo "Error description: " & Err.Description 
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
                    WScript.Echo "An unexpected program is used to run this script."
                    WScript.Echo "Only CScript.Exe or WScript.Exe can be used to run this script."
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
'* Output: strFile                 the output file name
'*             strCurrentUser    the name or cn of the user for the destination domain
'*             strPassword        the password for the user in the destination domain
'*             blnQuiet              specifies whether to suppress messages
'*             strServer             the destination server name
'*             intParseCmdLine     is set to one of CONST_ERROR, CONST_SHOW_USAGE, CONST_PROCEED.
'*
'********************************************************************

Private Function intParseCmdLine(strArgumentArray, strCurrentUser,strPassword, blnQuiet, strServer, strFile)

    ON ERROR RESUME NEXT

    Dim i, j, strFlag

    intParseCmdLine = CONST_PROCEED

    strFlag = strArgumentArray(0)
    If strFlag = "" then                    'No arguments have been received
	intParseCmdLine = CONST_SHOW_USAGE
        Exit Function
    End If

    'online help was requested
    If (strFlag="/help") OR (strFlag="/HELP") OR (strFlag="/h") OR (strFlag="\h") OR (strFlag="-h") _
        OR (strFlag = "\?") OR (strFlag = "/?") OR (strFlag = "?") OR (strFlag="h") Then
        intParseCmdLine = CONST_SHOW_USAGE
        Exit Function
    End If

    j = 0
    For i = 0 to UBound(strArgumentArray)
        strFlag = Left(strArgumentArray(i), InStr(1, strArgumentArray(i), ":")-1)
        If Err.Number Then            'An error occurs If there is no : in the string
            Err.Clear
            Select Case LCase(strArgumentArray(i))
                Case "/q" 
                  blnQuiet = True
                Case Else
                  WScript.Echo strArgumentArray(i) & " is NOT recognized as a valid input."
                  intParseCmdLine = CONST_ERROR
                  Exit Function
            End Select 'end processing args that have no params

        Else
            Select Case LCase(strFlag)
                Case "/u"
                    strCurrentUser = Right(strArgumentArray(i), Len(strArgumentArray(i))-3)
                Case "/w"
                    strPassword = Right(strArgumentArray(i), Len(strArgumentArray(i))-3)
                Case "/s"
                    strServer = Right(strArgumentArray(i), Len(strArgumentArray(i))-3)
                Case "/f"
	    strFile = Right(strArgumentArray(i), Len(strArgumentArray(i))-3)
                Case else
                  WScript.Echo strArgumentArray(i) & " is not recognized as a valid input."
                  intParseCmdLine = CONST_ERROR
                  Exit Function
            End Select
        End If
    Next

   If Len(strCurrentUser) = 0 Then
      WScript.Echo "User name not supplied using /u switch"	
      intParseCmdLine = CONST_SHOW_USAGE
   End If

   If Len(strServer) = 0 Then
      WScript.Echo "Server name not supplied using /s switch"	
      intParseCmdLine = CONST_SHOW_USAGE
   End If


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

    Wscript.echo ""
    Wscript.echo "SCHEMADIFF.VBS    Compares schema on two different domains."  & vbCRLF
    Wscript.echo "SCHEMADIFF.VBS /U:username /W:password /S:server [/F:filename] [/Q]"
    Wscript.echo " command line switches:"
    Wscript.echo "   /? /H /HELP    Displays this help message."
    Wscript.echo " command line parameters:"
    Wscript.echo "   /U:Username    Username for destination domain."
    Wscript.echo "   /W:Password    Password for destination domain."
    WScript.echo "   /S:Server      Name of destination domain controller."
    Wscript.echo "   /F:Filename    Valid filename for output file. (optional)" 
    WScript.echo ""
    Wscript.echo "EXAMPLES:"
    WScript.echo ""
    Wscript.echo "CSCRIPT SCHEMADIFF  /S:wombat.acme.com /U:Administrator /W:password"
    WScript.echo "   compares the schema for the domain of the currently logged on user with the"
    WScript.echo "   schema for the domain of which wombat.acme.com is a domain controller."
End Sub

'********************************************************************
'* Sub	SCHEMADIFF()
'* Purpose	compares two schemas
'* Input		strServer	Name of the domain controller
'*		strCurrentUser	User credentials
'*		strPassword	User password
'*		blnQuiet
'*		strFile		Name of output file
'*		
'* Output	None
'*		Displays differences
'*		Optionally writes output to strFile
'*
'********************************************************************
Sub SchemaDiff(strServer, strCurrentUser, strPassword, blnQuiet, strFile)

ON ERROR RESUME NEXT

	Dim objSrcSchema, strSrcSchema, strSrcServer, objSrcClass
 	Dim objRoot, objFileSystem, objFile, objProvider, blnMatch
 	Dim objDstSchema, strDstSchema, strDstServer, objDstClass
	Dim iSrcSchemaVersion, iDstSchemaVersion

	Dim objADO, objADOCommand, strFilter, strResults, strSearch, rsADO, iCount, arrTemp
	
	blnMatch = True

	If strFile = "" Then
	        	objFile = ""
    	Else
        		'Create a filesystem object
	        	Set objFileSystem = CreateObject("Scripting.FileSystemObject")
        		If Err.Number Then
            			WScript.Echo "Error 0x" & CStr(Hex(Err.Number)) & " opening a filesystem object."
            			If Err.Description <> "" Then WScript.Echo "Error description: " & Err.Description & "."
			Err.Clear            		
                	Else
       
	      	                 'Open the file for output
                		Set objFile = objFileSystem.OpenTextFile(strFile, 2, True)
		                If Err.Number then
                                		WScript.Echo "Error 0x" & CStr(Hex(Err.Number)) & " opening file " & strFile
                                		If Err.Description <> "" Then WScript.Echo "Error description: " & Err.Description & "." 
				Err.Clear
                        		End If
                	End If
	End If
	
	'bind to source schema
	Set objRoot = GetObject("LDAP://RootDSE")	
	If Err.Number then
		WScript.Echo "Error 0x" & CStr(Hex(Err.Number)) & " occurred in binding to Source RootDSE."
		If Err.Description <> "" Then WScript.Echo "Error description: " & Err.Description 
		Err.Clear
	        	Exit Sub
	End If
               strSrcSchema = objRoot.Get("schemaNamingContext")
               strSrcServer = objRoot.Get("dnsHostName") & "/"

               If blnQuiet Then WScript.Echo "Source Schema: " & strSrcSchema
               If blnQuiet Then WScript.Echo "Source Server: " & strSrcServer

	Set objSrcSchema = GetObject("LDAP://" & strSrcServer & strSrcSchema)
	If Err.Number then
		WScript.Echo "Error 0x" & CStr(Hex(Err.Number)) & " occurred in binding to Source Schema NC. "
		If Err.Description <> "" Then WScript.Echo "Error description: " & Err.Description 
		Err.Clear
        	Exit Sub
	End If

	iSrcSchemaVersion = objSrcSchema.objectVersion

	'bind to destination schema

	If Len(strServer) > 0 Then strServer = strServer & "/"

	Set objRoot = GetObject("LDAP://" & strServer & "RootDSE")
	If Err.Number then
		WScript.Echo "Error 0x" & CStr(Hex(Err.Number)) & " occurred in binding to Destination RootDSE."
		If Err.Description <> "" Then WScript.Echo "Error description: " & Err.Description 
		Err.Clear
	        Exit Sub
	End If
	
	strServer = objRoot.Get("dnsHostName") & "/"

	'get schema

	strDstSchema = objRoot.Get("schemaNamingContext")

                If blnQuiet Then WScript.Echo "Dest Schema: " & strDstSchema
                If blnQuiet Then WScript.Echo "Dest Server: " & strServer

	If strCurrentUser = "" Then            'no user credential is passed
        	Set objDstSchema = GetObject("LDAP://" & strServer & strDstSchema)
	Else
        	Set objProvider = GetObject("LDAP:")
		'Use user authentication

		Set objDstSchema = objProvider.OpenDsObject("LDAP://" & strServer & strDstSchema, strCurrentUser, strPassword,1)
	End If
	If Err.Number then
		WScript.Echo "Error 0x" & CStr(Hex(Err.Number)) & " occurred in binding to Dest Schema NC. "
		If Err.Description <> "" Then WScript.Echo "Error description: " & Err.Description 
		Err.Clear
        	Exit Sub
	End If

	iDstSchemaVersion = objDstSchema.objectVersion

 	'create ADO search for destination

	Set objADO = CreateObject("ADODB.Connection")
	If Err.Number then
		WScript.Echo "Error 0x" & CStr(Hex(Err.Number)) & " occurred in creating ADO Connection. "
		If Err.Description <> "" Then WScript.Echo "Error description: " & Err.Description 
		Err.Clear
        	Exit Sub
	End If

	objADO.Provider = "ADsDSOObject"

	If Len(strCurrentUser) > 0 Then 
		objADO.Properties("User ID") = strCurrentUser
		objADO.Properties("Password") = strPassword
		objADO.Properties("Encrypt Password") = True
	End If

	objADO.Open "STUFF"
	If Err.Number then
		WScript.Echo "Error 0x" & CStr(Hex(Err.Number)) & " occurred in opening ADO Connection. "
		If Err.Description <> "" Then WScript.Echo "Error description: " & Err.Description 
		Err.Clear
        	Exit Sub
	End If

	Set objADOCommand = CreateObject("ADODB.Command")
	If Err.Number then
		WScript.Echo "Error 0x" & CStr(Hex(Err.Number)) & " occurred in creating ADO Command. "
		If Err.Description <> "" Then WScript.Echo "Error description: " & Err.Description 
		Err.Clear
        	Exit Sub
	End If

	Set objADOCommand.ActiveConnection = objADO
	If Err.Number then
		WScript.Echo "Error 0x" & CStr(Hex(Err.Number)) & " occurred in opening ADO Connection. "
		If Err.Description <> "" Then WScript.Echo "Error description: " & Err.Description 
		Err.Clear
        	Exit Sub
	End If


	objADOCommand.Properties("Timeout")=0
	objADOCommand.Properties("Time Limit")=0
	objADOCommand.Properties("Page Size")=100
	objADOCommand.Properties("Chase Referrals") = True

	WScript.Echo "Src Schema Version: " & iSrcSchemaVersion
	WScript.Echo "Dst Schema Version: " & iDstSchemaVersion
	
	If isObject(objFile) Then 
		objFile.WriteLine  "Src Schema Version: " & iSrcSchemaVersion
		objFile.WriteLine  "Dst Schema Version: " & iDstSchemaVersion
	End If

                objSrcSchema.Filter=Array("classSchema")
                For Each objSrcClass in objSrcSchema
		If Err.Number then
			WScript.Echo "Error 0x" & CStr(Hex(Err.Number)) & " occurred enumerating classes. "
			If Err.Description <> "" Then WScript.Echo "Error description: " & Err.Description 
			Err.Clear
	        		Exit Sub
		End If

		
		If blnQuiet Then WScript.Echo "Class: " & objSrcClass.ldapDisplayName
		If isObject(objFile) Then objFile.WriteLine "Class: " & objSrcClass.ldapDisplayName

		strFilter = "(&(objectClass=classSchema)(ldapDisplayName=" & objSrcClass.ldapDisplayName & "))"
		strResults = "cn,ldapDisplayName,governsID,systemMayContain,systemMustContain"
		strSearch = "<LDAP://" & strServer & strDstSchema & ">"

		objADOCommand.CommandText = strSearch & ";" & strFilter & ";" & strResults & ";subtree"
                                
		Set rsADO = objADOCommand.Execute

		If Err.Number then
			WScript.Echo "Error 0x" & CStr(Hex(Err.Number)) & " occurred in executing ADO Query. "
			If Err.Description <> "" Then WScript.Echo "Error description: " & Err.Description 
			Err.Clear
	        		Exit Sub
		End If

		
		
		If rsADO.RecordCount > 0 Then

			While Not rsADO.EOF
				If blnQuiet Then WScript.Echo rsADO.Fields(1).Value & " found"
				If blnQuiet Then WScript.Echo "Src OID: " & objSrcClass.governsID

                                                                If StrComp(rsADO.Fields(2).Value,objSrcClass.governsID) <> 0 Then
					WScript.Echo "Dst OID: " & rsADO.Fields(2).Value
					blnMatch=False
					If isObject(objFile) Then objFile.WriteLine "Dst OID: " & rsADO.Fields(2).Value
				Else
					If blnQuiet Then WScript.Echo "Src and Dst OID match"
				End If

                                                                If Compare2Arrays(objSrcClass.systemMayContain,rsADO.Fields(3).Value) <>CONST_PROCEED Then
					WScript.Echo "Dst systemMayContain mismatch "
					blnMatch=False
					If isObject(objFile) Then  objFile.WriteLine "Dst systemMayContain mismatch"

				Else
					if blnQuiet Then WScript.Echo "systemMayContain OK"
				End If

				If Compare2Arrays(objSrcClass.systemMustContain,rsADO.Fields(4).Value) <>CONST_PROCEED Then
					WScript.Echo "Dst systemMustContain mismatch"
					blnMatch=False
					If isObject(objFile) Then  objFile.WriteLine "Dst systemMustContain mismatch"
				Else
					if blnQuiet Then WScript.Echo "systemMustContain OK"
				End If

				rsADO.MoveNext
			Wend
		

		Else
			WScript.Echo "Missing Class " & objSrcClass.ldapDisplayName
			blnMatch=False
			If isObject(objFile) Then objFile.WriteLine "Missing Class " & objSrcClass.ldapDisplayName
			If Err.Number then
				WScript.Echo "Error 0x" & CStr(Hex(Err.Number)) & " occurred in missing class. "
				If Err.Description <> "" Then WScript.Echo "Error description: " & Err.Description 
				Err.Clear
			End If
		End If
	Next
	
	'now search for attributes
	WScript.Echo "Searching attributes"

                objSrcSchema.Filter=Array("attributeSchema")
	
                For Each objSrcClass in objSrcSchema
		If Err.Number then
			WScript.Echo "Error 0x" & CStr(Hex(Err.Number)) & " occurred enumerating attributes. "
			If Err.Description <> "" Then WScript.Echo "Error description: " & Err.Description 
			Err.Clear
	        		Exit Sub
		End If
		If blnQuiet Then WScript.Echo "Attribute: " & objSrcClass.ldapDisplayName
		If isObject(objFile) Then objFile.WriteLine "Class: " & objSrcClass.ldapDisplayName

		strFilter = "(&(objectCategory=attributeSchema)(ldapDisplayName=" & objSrcClass.ldapDisplayName & "))"
		strResults = "cn,ldapDisplayName,attributeID,attributeSyntax,rangeLower,rangeUpper"
		strSearch = "<LDAP://" & strServer & strDstSchema & ">"
		objADOCommand.CommandText = strSearch & ";" & strFilter & ";" & strResults & ";subtree"

		
		Set rsADO = objADOCommand.Execute
		If Err.Number then
			WScript.Echo "Error 0x" & CStr(Hex(Err.Number)) & " occurred in executing ADO Query. "
			If Err.Description <> "" Then WScript.Echo "Error description: " & Err.Description 
			Err.Clear
	        		Exit Sub
		End If

		
		iCount = 0
		If rsADO.RecordCount > 0 Then

			While Not rsADO.EOF

				If blnQuiet Then WScript.Echo rsADO.Fields(1).Value & " found"
				If blnQuiet Then WScript.Echo "Attribute OID: " & objSrcClass.attributeID

				If StrComp(objSrcClass.attributeID,rsADO.Fields(2).Value) <> 0 Then	
					WScript.Echo "Dst OID: " & rsADO.Fields(2).Value
					blnMatch=False
					If isObject(objFile) Then objFile.WriteLine "Dst OID: " & rsADO.Fields(2).Value
				Else
					If blnQuiet Then WScript.Echo "Src and Dst OID match"
				End If

				if StrComp(objSrcClass.attributeSyntax ,rsADO.Fields(3).Value) <> 0 Then	
					WScript.Echo "Dst syntax: " & sADO.Fields(3).Value
					blnMatch=False
					If isObject(objFile) Then objFile.WriteLine "Dst syntax: " & sADO.Fields(3).Value
				Else
					If blnQuiet Then WScript.Echo "Attribute Syntax OK"
				End If

				if objSrcClass.rangeLower <> rsADO.Fields(4).Value Then	
					WScript.Echo "Dst Range Lower " & rsADO.Fields(4).Value
					blnMatch=False
					If isObject(objFile) Then objFile.WriteLine "Dst Range Lower " & rsADO.Fields(4).Value
				Else
					If blnQuiet Then WScript.Echo "Range Lower OK"
				End If
				
				if objSrcClass.rangeUpper <> rsADO.Fields(5).Value Then	
					WScript.Echo "Dst Range Upper " & rsADO.Fields(5).Value
					blnMatch=False
					If isObject(objFile) Then objFile.WriteLine "Dst Range Upper " & rsADO.Fields(5).Value
				Else
					If blnQuiet Then WScript.Echo "Range Upper OK"
				End If


				rsADO.MoveNext
			Wend
		
			
                	
		Else
			WScript.Echo "Missing Attribute " & objSrcClass.ldapDisplayName
			blnMatch=False
			If isObject(objFile) Then objFile.WriteLine "Missing Attribute " & objSrcClass.ldapDisplayName

		End If
	Next

	If blnMatch Then 
                       	WScript.Echo "** Schema is OK"
		If isObject(objFile) Then objFile.WriteLine "** Schema is OK"
	Else
		WScript.Echo "** Schema Mismatch"
		If isObject(objFile) Then objFile.WriteLine "** Schema Mismatch"
	End If

	If isObject(objFile) Then objFile.Close
	
End Sub


'********************************************************************
'* Function	Compare2Arrays()
'* Input		SrcArray
'*		DstArray
'* Output		If DstArray contains all the same elements as SrcArray Compare2Arrays 
'*		returns CONST_PROCEED else returns CONST_STRING_NOT_FOUND
'*
'********************************************************************
Function Compare2Arrays(SrcArray, DstArray)
ON ERROR RESUME NEXT

Dim i, j, k, bFound

Compare2Arrays = CONST_STRING_NOT_FOUND


If IsEmpty(SrcArray) And IsEmpty(DstArray) Then 
	Compare2Arrays = CONST_PROCEED
	Exit Function	
End If


If IsEmpty(SrcArray) And IsNull(DstArray) Then 
	Compare2Arrays = CONST_PROCEED
	Exit Function	
End If


If IsNull(SrcArray) And IsNull(DstArray) Then 
	Compare2Arrays = CONST_PROCEED
	Exit Function	
End If


If IsNull(SrcArray) And IsEmpty(DstArray) Then 
	Compare2Arrays = CONST_PROCEED
	Exit Function	
End If


If (Not IsArray(SrcArray)) And (Not IsArray(DstArray)) Then
	If StrComp(SrcArray, DstArray) = 0 Then Compare2Arrays = CONST_PROCEED
	Exit Function
End If

If (Not IsArray(SrcArray)) And (IsArray(DstArray)) Then
	If StrComp(SrcArray, DstArray(UBound(DstArray))) = 0 Then 	Compare2Arrays = CONST_PROCEED
	Exit Function
End If

If (IsArray(SrcArray)) And (Not IsArray(DstArray)) Then
	If StrComp(DstArray, SrcArray(UBound(SrcArray))) = 0 Then 	Compare2Arrays = CONST_PROCEED
	Exit Function
End If

If UBound(SrcArray) <> UBound(DstArray) Then 
	Exit Function
End If

For i = LBound(SrcArray) to UBound(SrcArray)
	bFound=False
	For j = LBound(DstArray) to UBound(DstArray)
		If StrComp(LCase(SrcArray(i)) ,LCase(DstArray(j))) = 0  Then bFound = True
	Next
	if bFound = False Then Exit Function
Next

Compare2Arrays = CONST_PROCEED

End Function



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

Function intSearchArray(ByVal strTarget, ByVal strArray)

ON ERROR RESUME NEXT

    Dim i, j

    intSearchArray = CONST_STRING_NOT_FOUND

    If Not IsArray(strArray) Then
        WScript.Echo "Argument is not an array!"
        Exit Function
    End If

    strTarget = LCase(strTarget)
    For i = 0 To UBound(strArray)
        j = InStr(1, strArray(i), strTarget, 1)
        If j > 0 Then 
           intSearchArray = i
        End If
    Next
End Function

'********************************************************************
'*
'* Sub Print()
'* Purpose:   Prints a message on screen If blnQuiet = False.
'* Input:     strMessage    the string to print
'* Output:    strMessage is printed on screen If blnQuiet = False.
'*
'********************************************************************

Sub Print(ByRef strMessage)
    If Not blnQuiet then
        Wscript.Echo  strMessage
    End If
End Sub

'********************************************************************
'*
'* Funcion sprintf()
'* Purpose:   formats a string simlar to C runtime sprintf function
'* Input:     VarArg    variable length array of strings
'* Output:    formatted string
'*
'********************************************************************
Public Function sprintf(VarArg())
Dim iIndex
Dim iArg
Dim iCountArgs
Dim sTemp
Dim cChar
Dim cNextChar
Dim bFound
Dim sVal

iArg = 1
iCountArgs = UBound(VarArg) + 1

For iIndex = 1 to Len(VarArg(0))
  cChar = Mid(VarArg(0),iIndex,1)
  Select Case cChar
  Case "%"
    bFound = False
    sVal = 0
    Do While bFound=False
      cNextChar=Mid(VarArg(0),iIndex+1,1)
      iIndex=iIndex+1
        Select Case cNextChar
        Case "d","s"
            bFound=True
        Case Else
            If sVal > 0 Then sVal = sVal * 10
            sVal = sVal + cNextChar
        End Select
      Loop
      If iArg < iCountArgs Then
        If Len(VarArg(iArg)) > sVal Then 
           sTemp = sTemp & Left(VarArg(iArg),sVal)
        Else
           sTemp = sTemp & VarArg(iArg) & Space(sVal-Len(VarArg(iArg))) 
        End If 
      End If
      iArg = iArg +1
  Case Else
  sTemp = sTemp & cChar
  End Select
Next
sprintf = sTemp
End Function


'********************************************************************
'*                                                                  *
'*                           End of File                      *
'*                                                                  *
'********************************************************************