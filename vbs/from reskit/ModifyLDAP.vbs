
'********************************************************************
'*
'* File:        MODIFYLDAP.VBS
'* Created:     January 1999
'* Version:     1.0
'*
'* Main Function: Modifies LDAP Parameters.
'* Usage: MODIFYLDAP.VBS /A|T|C|D|M|P|R [/O:policy] [property1:propertyvalue1]
'*        [property2:propertyvalue2 ...] [/U:username] [/W:password] [/S:server|site] [/Q]
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
CONST CONST_MODIFY                      = 5
CONST CONST_CREATE                      = 6
CONST CONST_DELETE                      = 7
CONST CONST_ASSIGN                      = 8
CONST CONST_PRINT                       = 9
CONST CONST_SITE                        = 10
CONST CONST_REMOVE                      = 11


CONST ADS_OBJECT_NOTFOUND               = &H80072030
CONST ADS_OBJECT_EXISTS                 = &H80071392
CONST ADS_PROPERTY_CLEAR                = 1
CONST ADS_PROPERTY_UPDATE               = 2
CONST ADS_PROPERTY_APPEND               = 3
CONST ADS_PROPERTY_DELETE               = 4

'Declare variables
Dim strDomain, strFile, strCurrentUser, strPassword, strPolicy, strServer, blnQuiet, intOpMode, i
ReDim strArgumentArray(0), strPropertyArray(0), strPropertyValueArray(0)

'Initialize variables
intOpMode = 0
blnQuiet = False
strPolicy = ""
strFile = ""
strCurrentUser = ""
strPassword = ""
strServer = ""
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
            "1. Using ""CScript MODIFYLDAP.vbs arguments"" for Windows 95/98 or" & vbCRLF & _
            "2. Changing the default Windows Scripting Host setting to CScript" & vbCRLF & _
            "    using ""CScript //H:CScript //S"" and running the script using" & vbCRLF & _
            "    ""MODIFYLDAP.vbs arguments"" for Windows NT."
        WScript.Quit
    Case Else
        WScript.Quit
End Select

'Parse the command line
intOpMode = intParseCmdLine(strArgumentArray, strPolicy, strFile, strCurrentUser,_
        strPassword, blnQuiet, strServer, strPropertyArray, strPropertyValueArray)
If Err.Number Then
    Print "Error 0x" & CStr(Hex(Err.Number)) & " occurred in parsing the command line."
    If Err.Description <> "" Then  Print "Error description: " & Err.Description & "."
    WScript.quit
End If

Select Case intOpMode
    Case CONST_SHOW_USAGE
        Call ShowUsage()
    Case CONST_CREATE
        Call CREATELDAP(strPolicy, strCurrentUser, strPassword, strPropertyArray, strPropertyValueArray, blnQuiet)
    Case CONST_MODIFY
        Call MODIFYLDAP(strPolicy, strFile, strCurrentUser,_
             strPassword, blnQuiet, strPropertyArray, strPropertyValueArray)
    Case CONST_DELETE
        Call DELETELDAP(strPolicy, strCurrentUser, strPassword, blnQuiet)
    Case CONST_ASSIGN
        Call ASSIGNLDAP(strPolicy, strServer, strCurrentUser, strPassword, blnQuiet)
    Case CONST_PRINT
        Call PRINTLDAP(strPolicy, strFile, strCurrentUser, strPassword, blnQuiet)
    Case CONST_REMOVE
        Call REMOVELDAP(strServer, strCurrentUser, strPassword, blnQuiet)
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
        Print "Error 0x" & CStr(Hex(Err.Number)) & " occurred."
        If Err.Description <> "" Then Print "Error description: " & Err.Description & "."
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
'* Output:  strPolicy           the name of the policy object
'*          strFile             the input file name including the path
'*          strCurrentUser      the name or cn of the current user
'*          strPassword         the current user password
'*          blnQuiet            specifies whether to suppress messages
'*          strServer           server name
'*          strPropertyArray    an array of ldap parameters
'*          strPropertyValueArray    an array of the corresponding ldap parameter values
'*          intParseCmdLine     is set to one of CONST_ERROR, CONST_SHOW_USAGE, CONST_PROCEED.
'*
'********************************************************************

Private Function intParseCmdLine(strArgumentArray, strPolicy, strFile, strCurrentUser,_
        strPassword, blnQuiet, strServer, strPropertyArray, strPropertyValueArray)

    ON ERROR RESUME NEXT

    Dim i, j, strFlag

    intParseCmdLine = CONST_ERROR

    strFlag = strArgumentArray(0)
    If strFlag = "" then                    'No arguments have been received
        Print "Arguments are required."
        intParseCmdLine = CONST_ERROR
        Exit Function
    End If

    'online help was requested
    If (strFlag="help") OR (strFlag="/h") OR (strFlag="\h") OR (strFlag="-h") _
        OR (strFlag = "\?") OR (strFlag = "/?") OR (strFlag = "?") OR (strFlag="h") Then
        intParseCmdLine = CONST_SHOW_USAGE
        Exit Function
    End If

    j = 0
    For i = 0 to UBound(strArgumentArray)
        strFlag = Left(strArgumentArray(i), InStr(1, strArgumentArray(i), ":")-1)
        If Err.Number Then            'An error occurs if there is no : in the string
            Err.Clear
            Select Case LCase(strArgumentArray(i))
                Case "/q" 
                  blnQuiet = True
                Case "/a"
                  intParseCmdLine = CONST_ASSIGN
                Case "/d"
                  intParseCmdLine = CONST_DELETE
                Case "/m"
                  intParseCmdLine = CONST_MODIFY
                Case "/c"
                  intParseCmdLine = CONST_CREATE
                Case "/p"
                  intParseCmdLine = CONST_PRINT
                Case "/r"
                  intParseCmdLine = CONST_REMOVE
                Case Else
                  Print strArgumentArray(i) & " is not recognized as a valid input."
                  intParseCmdLine = CONST_ERROR
                  Exit Function
            End Select 'end processing args that have no params

        Else
            Select Case LCase(strFlag)
                Case "/o" 
                    strPolicy = Right(strArgumentArray(i), Len(strArgumentArray(i))-3)
                Case "/f"
                    strFile = Right(strArgumentArray(i), Len(strArgumentArray(i))-3)
                Case "/u"
                    strCurrentUser = Right(strArgumentArray(i), Len(strArgumentArray(i))-3)
                Case "/w"
                    strPassword = Right(strArgumentArray(i), Len(strArgumentArray(i))-3)
                Case "/s"
                    strServer = Right(strArgumentArray(i), Len(strArgumentArray(i))-3)
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


    If ((strPolicy = "") And (intParseCmdLine <> CONST_REMOVE)) Then 
          print "The name of the policy object is missing."
          intParseCmdLine = CONST_ERROR
          Exit Function
    End If

    If ((strServer = "") And ((intParseCmdLine = CONST_REMOVE) Or (intParseCmdLine = CONST_ASSIGN))) Then 
          print "The name of the site or server is missing."
          intParseCmdLine = CONST_ERROR
          Exit Function
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
    Wscript.echo "Modifies LDAP policies."  & vbCRLF
    Wscript.echo "MODIFYLDAP.VBS [/A|T|C|D|M|P|R] /O:Policy [/F:filename] [property1:propertyvalue1]" 
    Wscript.echo "[property2:propertyvalue2...] [/U:username] [/W:password] [/S:server] [/Q]"
    Wscript.echo " command line switches:"
    WScript.echo "   /C             Create a new query policy."
    Wscript.echo "   /D             Delete an existing query policy."
    Wscript.echo "   /M             Modify an existing query policy."
    Wscript.echo "   /A             Assign a policy to a site or server."
    Wscript.echo "   /R             Remove a policy from a site or server."
    Wscript.echo "   /P             Display the LDAP Admin limits for a query policy."
    Wscript.echo "   /Q             Quiet mode." 
    Wscript.echo "   /? /H /HELP    Displays this help message."
    Wscript.echo "   property1:propertyvalue1...property[i], propertyvalue[i]"
    WScript.echo "                  Name and value of ldap parameters."
    Wscript.echo " command line parameters:"
    Wscript.echo "   /F:filename   valid filename." 
    Wscript.echo "   /U:username    Username."
    Wscript.echo "   /W:password    Password."
    WScript.echo "   /S:server      Name of domain controller or site."
    WScript.echo "   /O:policy      Name of Policy object."
    Wscript.echo "EXAMPLES:"
    Wscript.echo "MODIFYLDAP.VBS /C /O:NewPolicy ConnectionTimeOut:1000"
    Wscript.echo "   Creates a policy named NewPolicy with ConnectionTimeOut=1000"
    WScript.echo "   and defaults for remainder." 
    Wscript.echo "MODIFYLDAP.VBS /D /O:NewPolicy"
    Wscript.echo "   Deletes the LDAP Policy NewPolicy."
    Wscript.echo "MODIFYLDAP.VBS /M /O:NewPolicy InitRecvTimeout:200"
    Wscript.echo "   Modifies the LDAP Policy NewPolicy setting InitRecvTimeout value to 200."
    Wscript.echo "MODIFYLDAP.VBS /A /O:NewPolicy /S:Wombat"
    Wscript.echo "   Assigns the LDAP Policy NewPolicy to the server called Wombat."
    Wscript.echo "MODIFYLDAP.VBS /R wombat"
    Wscript.echo "   Removes the LDAP Policy associated with the site Timbucktoo."
    Wscript.echo "MODIFYLDAP.VBS /P /O:NewPolicy"
    Wscript.echo "   Displays the current settings for the LDAP Policy NewPolicy."


End Sub

'********************************************************************
'*
'* Sub MODIFYLDAP()
'* Purpose: Modifies an existing policy object.
'* Input:   strPolicy           the name of the policy object
'*          strFile             the input file name including the path
'*          strCurrentUser      the name or cn of the current user
'*          strPassword         the current user password
'*          blnQuiet            specifies whether to suppress messages
'*          strPropertyArray    an array of ldap properties names
'*          strPropertyValueArray    an array of the corresponding ldap property values
'* Output:  None
'*
'********************************************************************

Sub MODIFYLDAP(strPolicy, strFile, strCurrentUser,_
    strPassword, blnQuiet, strPropertyArray, strPropertyValueArray)

    ON ERROR RESUME NEXT

    Dim objDomain, strUser, objPolicy, i, j, objFileSystem, objInputFile, strInput
    Dim blnResult, strConfig, objProvider, arrLimits

    Set objDomain = GetObject("LDAP://RootDSE")
    strConfig = objDomain.Get("configurationNamingContext")
    Set objDomain=Nothing
   
    If Err.Number then
	Print "Error 0x" & CStr(Hex(Err.Number)) & " occurred in getting config nc."
	If Err.Description <> "" Then Print "Error description: " & Err.Description & "."
    	Err.Clear
        Exit Sub
    End If

    If strCurrentUser = "" Then            'no user credential is passed
        Set objDomain = GetObject("LDAP://CN=Query-Policies,CN=Directory Service,CN=Windows NT,CN=Services," & strConfig)
    Else
        Set objProvider = GetObject("LDAP:")
		'Use user authentication
        Set objDomain = objProvider.OpenDsObject("LDAP://CN=Query-Policies,CN=Directory Service,CN=Windows NT,CN=Services," _
                           & strConfig,strCurrentUser,strPassword,1)
    End If
    If Err.Number then
	Print "Error 0x" & CStr(Hex(Err.Number)) & " occurred in getting object " _
		& strPolicy & "."
	If Err.Description <> "" Then Print "Error description: " & Err.Description & "."
	Err.Clear
        Exit Sub
    End If

    'read the ldap admin limits into an array
    Set objPolicy = objDomain.GetObject("queryPolicy","cn=" &  strPolicy)
    If Err.Number then
	Print "Error 0x" & CStr(Hex(Err.Number)) & " occurred in getting object " _
		& strPolicy & "."
	If Err.Description <> "" Then Print "Error description: " & Err.Description & "."
	Err.Clear
        Exit Sub
    End If
 
    
    arrLimits = objPolicy.ldapAdminLimits

    For j = LBound(StrPropertyArray) to UBound(StrPropertyArray)
       i = intSearchArray(strPropertyArray(j), arrLimits)
       if i <> -1 Then arrLimits(i) = strPropertyArray(j) & "=" & strPropertyValueArray(j)
    Next

    If i <> -1 Then
       objPolicy.PutEx ADS_PROPERTY_UPDATE, "ldapAdminLimits", arrLimits
       objPolicy.Setinfo
       If Err.Number then
		Print "Error 0x" & CStr(Hex(Err.Number)) & " occurred in setting object " & strPolicy & "."
		If Err.Description <> "" Then	Print "Error description: " & Err.Description & "."
		Err.Clear
        	Exit Sub
	Else
        	Print strPolicy & " updated"
	End If

    Else
	Print strPolicy & " not updated due to incorrect LDAP Query parameters"
    End If

    Set objPolicy = Nothing
    Set objDomain = Nothing

End Sub



'********************************************************************
'*
'* Sub PRINTLDAP()
'* Purpose: Prints the LDAP parameters of a LDAP policy object.
'* Input:   strPolicy           the ADsPath of the domain
'*          strFile             the file name to print the data
'*          strCurrentUser      the name or cn of the current user
'*          strPassword         the current user password
'*          blnQuiet            specifies whether to suppress messages
'* Output:  None
'*
'********************************************************************

Sub PRINTLDAP(strPolicy, strFile, strCurrentUser, strPassword, blnQuiet)

    ON ERROR RESUME NEXT

    Dim  objDomain, objProvider, strConfig, objPolicy, i, objFileSystem, objOutputFile, strInput
    Dim blnResult, arrLimits


    If strFile <> "" Then
        'Create a filesystem object
        set objFileSystem = CreateObject("Scripting.FileSystemObject")
        If Err.Number Then
            Print "Error 0x" & CStr(Hex(Err.Number)) & " occurred in creating a filesystem object."
            If Err.Description <> "" Then Print "Error description: " & Err.Description & "."
            Exit Sub
        End If

        'Opens a file for output
        set objOutputFile = objFileSystem.OpenTextFile(strFile,2,True)
        If Err.Number Then
            Print "Error 0x" & CStr(Hex(Err.Number)) & " occurred in opening file " & strFile
            If Err.Description <> "" Then Print "Error description: " & Err.Description & "."
            Exit Sub
        End If
    End If


    Set objDomain = GetObject("LDAP://RootDSE")
    strConfig = objDomain.Get("configurationNamingContext")
    Set objDomain=Nothing
   
    If Err.Number then
	Print "Error 0x" & CStr(Hex(Err.Number)) & " occurred in getting configuration container."
	If Err.Description <> "" Then Print "Error description: " & Err.Description & "."
	Err.Clear
        Exit Sub
    End If

    If strCurrentUser = "" Then            'no user credential is passed
        Set objDomain = GetObject("LDAP://CN=Query-Policies,CN=Directory Service,CN=Windows NT,CN=Services," & strConfig)
    Else
        Set objProvider = GetObject("LDAP:")
		'Use user authentication
        Set objDomain = objProvider.OpenDsObject("LDAP://CN=Query-Policies,CN=Directory Service,CN=Windows NT,CN=Services," _
                           & strConfig,strCurrentUser,strPassword,1)
    End If
    If Err.Number then
	Print "Error 0x" & CStr(Hex(Err.Number)) & " occurred in getting object " & strPolicy & "."
	If Err.Description <> "" Then Print "Error description: " & Err.Description & "."
	Err.Clear
        Exit Sub
    End If

'*** Query all not supported yet ***
'
'    If blnAll Then        'retrieve all policy objects
'	objDomain.Filter= Array("queryPolicy")
'        For Each objPolicy in ObjDomain
'            print objPolicy.cn
'            If strFile <> "" Then objOutputFile.WriteLine objPolicy.cn
'            arrLimits = objPolicy.ldapAdminLimits
'            For i = LBound(arrLimits) to UBound(arrLimits)
'               print arrLimits(i)
'               If strFile <> "" Then objOutputFile.WriteLine arrLimits(i)
'            Next
'        Next
'    Else                  'retrieve specified policy object

    Set objPolicy = objDomain.GetObject("queryPolicy","cn=" &  strPolicy)
    If Err.Number then
	Print "Error 0x" & CStr(Hex(Err.Number)) & " occurred in getting object " & strPolicy & "."
	If Err.Description <> "" Then Print "Error description: " & Err.Description & "."
        Err.Clear
        Exit Sub
    End If

    print "LDAP Settings for " & objPolicy.cn
    If strFile <> "" Then objOutputFile.WriteLine objPolicy.cn
    arrLimits = objPolicy.ldapAdminLimits
    For i = LBound(arrLimits) to UBound(arrLimits)
         print arrLimits(i)
         If strFile <> "" Then objOutputFile.WriteLine arrLimits(i)
    Next

    Set objPolicy = Nothing
    Set objDomain = Nothing

    If strFile <> "" Then 
            objOutputFile.Close
    End If

End Sub


'********************************************************************
'*
'* Sub ASSIGNLDAP()
'* Purpose: ASSIGNS LDAP policy object to a domain controller
'* Input:   strPolicy           the ADsPath of the domain
'*          strServer           The name of the domain controller
'*          strCurrentUser      the name or cn of the current user
'*          strPassword         the current user password
'*          blnQuiet            specifies whether to suppress messages
'* Output:  None
'*
'********************************************************************

Sub ASSIGNLDAP(strPolicy, strServer, strCurrentUser, strPassword, blnQuiet)
    ON ERROR RESUME NEXT

    Dim  objDomain, strConfig, objPolicy, strADSPath, objServer, blnSite
    Dim  objProvider, strCriteria, strPolicyPath, strClass, objNTDSSettings


    Set objDomain = GetObject("LDAP://RootDSE")
    strConfig = objDomain.Get("configurationNamingContext")
    Set objDomain=Nothing
   
    If Err.Number then
	Print "Error 0x" & CStr(Hex(Err.Number)) & " occurred in getting configuration container."
	If Err.Description <> "" Then	Print "Error description: " & Err.Description & "."
	Err.Clear
        Exit Sub
    End If

    If strCurrentUser = "" Then            'no user credential is passed
        Set objDomain = GetObject("LDAP://CN=Query-Policies,CN=Directory Service,CN=Windows NT,CN=Services," & strConfig)
    Else
        Set objProvider = GetObject("LDAP:")
		'Use user authentication
        Set objDomain = objProvider.OpenDsObject("LDAP://CN=Query-Policies,CN=Directory Service,CN=Windows NT,CN=Services," _
                           & strConfig,strCurrentUser,strPassword,1)
    End If
    If Err.Number then
	Print "Error 0x" & CStr(Hex(Err.Number)) & " occurred in getting query policy container."
	If Err.Description <> "" Then Print "Error description: " & Err.Description & "."
	Err.Clear
        Exit Sub
    End If

    Set objPolicy = objDomain.GetObject("queryPolicy","cn=" &  strPolicy)
    If Err.Number then
	Print "Error 0x" & CStr(Hex(Err.Number)) & " occurred in getting object " & strPolicy & "."
	If Err.Description <> "" Then Print "Error description: " & Err.Description & "."
	Err.Clear
        Exit Sub
    End If

    strPolicyPath = objPolicy.distinguishedName
    
    Set objPolicy = Nothing
    Set objDomain = Nothing
    Set ObjProvider = Nothing

    strCriteria = "(&(|(objectClass=site)(objectClass=server))(cn=" & strServer & "))"
        
    if blnSearchForServer(strServer, strConfig, strCurrentUser, strPassword, strCriteria, strADSPath, strClass) Then
        For i = LBound(strClass) to UBound(strClass)
      	   If InStr(1, strClass(i), "site",1) then 
	        blnSite = True
	        Exit For
           End If  
        Next
         
    	If strCurrentUser = "" Then            'no user credential is passed
        	Set objServer = GetObject(strADSPath)
	Else
        	Set objProvider = GetObject("LDAP:")
		'Use user authentication
	        Set objServer = objProvider.OpenDsObject(strADSPath,strCurrentUser,strPassword,1)
	End If
    	If Err.Number then
		Print "Error 0x" & CStr(Hex(Err.Number)) & " occurred in getting object " & strADSPath & "."
		If Err.Description <> "" Then Print "Error description: " & Err.Description & "."
		Err.Clear
	        Exit Sub
    	End If

        If blnSite Then
        	Set objNTDSSettings = objServer.GetObject("nTDSSiteSettings", "cn=NTDS Site Settings")
        Else
         	Set objNTDSSettings = objServer.GetObject("nTDSDSA", "cn=NTDS Settings")
        End If
        If Err.Number then
		Print "Error 0x" & CStr(Hex(Err.Number)) & " occurred in getting NTDS object " & strADSPath & "."
		If Err.Description <> "" Then Print "Error description: " & Err.Description & "."
		Err.Clear
	        Exit Sub
    	End If

        objNTDSSettings.Put "queryPolicyObject", strPolicyPath
        objNTDSSettings.SetInfo

        If Err.Number then
		Print "Error 0x" & CStr(Hex(Err.Number)) & " occurred in setting NTDSDSDA object."
		If Err.Description <> "" Then Print "Error description: " & Err.Description & "."
		Err.Clear
                Print "Server " & strServer & " was not updated"
	        Exit Sub
        Else
           Print "Server " & strServer & " was updated"
    	End If
    Else
       Print "Server " & strServer & " was not found"
    End If

    Set objNTDSSettings = Nothing
    Set objServer = Nothing
    
End Sub


'********************************************************************
'*
'* Sub CREATELDAP()
'* Purpose: Creates a new LDAP policy object.
'* Input:   strPolicy           the name of the policy object
'*          strCurrentUser      the name or cn of the current user
'*          strPassword         the current user password
'*          strPropertyArray    array of ldap parameters
'*          strPropertyValueArray array of ldap parameter values
'*          blnQuiet            specifies whether to suppress messages
'* Output:  None
'*
'********************************************************************

Sub CREATELDAP(strPolicy, strCurrentUser, strPassword, strPropertyArray, strPropertyValueArray, blnQuiet)

    ON ERROR RESUME NEXT

    Dim  objDomain, objProvider, strConfig, objPolicy, i, j, arrLimits

    Redim arrLimits(11)

    'these are the current defaults

    arrLimits(0) = "MaxConnections=1000"
    arrLimits(1) = "MaxDatagramRecv=1024"
    arrLimits(2) = "MaxPoolThreads=4"
    arrLimits(3) = "MaxResultSetSize=262144"
    arrLimits(4) = "MaxTempTableSize=10000"
    arrLimits(5) = "MaxQueryDuration=120"
    arrLimits(6) = "MaxPageSize=1000"
    arrLimits(7) = "MaxNotificationPerConn=5"
    arrLimits(8) = "MaxActiveQueries=20"
    arrLimits(9) = "MaxConnIdleTime=900"
    arrLimits(10) = "AllowDeepNonIndexSearch=False"
    arrLimits(11) = "InitRecvTimeout=120"

        
    If Not IsNull(strPropertyArray) Then 
      For j = LBound(StrPropertyArray) to UBound(StrPropertyArray)
         i = intSearchArray(strPropertyArray(j), arrLimits)
         If i <> -1 Then 
             arrLimits(i) = strPropertyArray(j) & "=" & strPropertyValueArray(j)
         Else
             print "Invalid LDAP parameter: " & strPropertyArray(j) & " ,using correct default values"
         End If
      Next
    End If


    Set objDomain = GetObject("LDAP://RootDSE")
    strConfig = objDomain.Get("configurationNamingContext")
    Set objDomain=Nothing
   
    If Err.Number then
	Print "Error 0x" & CStr(Hex(Err.Number)) & " occurred in getting configuration container."
	If Err.Description <> "" Then Print "Error description: " & Err.Description & "."
	Err.Clear
        Exit Sub
    End If

    If strCurrentUser = "" Then            'no user credential is passed
        Set objDomain = GetObject("LDAP://CN=Query-Policies,CN=Directory Service,CN=Windows NT,CN=Services," & strConfig)
    Else
        Set objProvider = GetObject("LDAP:")
		'Use user authentication
        Set objDomain = objProvider.OpenDsObject("LDAP://CN=Query-Policies,CN=Directory Service,CN=Windows NT,CN=Services," _
                           & strConfig,strCurrentUser,strPassword,1)
    End If
    If Err.Number then
	Print "Error 0x" & CStr(Hex(Err.Number)) & " occurred in getting query policy container. "
	If Err.Description <> "" Then Print "Error description: " & Err.Description & "."
	Err.Clear
        Exit Sub
    End If

    Set objPolicy = objDomain.Create("queryPolicy","cn=" &  strPolicy)
    objPolicy.PutEx  ADS_PROPERTY_UPDATE, "lDAPAdminLimits", arrLimits
    objPolicy.SetInfo
    If Err.Number then
	Print "Error 0x" & CStr(Hex(Err.Number)) & " occurred in creating " & strPolicy & "."
	If Err.Description <> "" Then Print "Error description: " & Err.Description & "."
	Err.Clear
        Exit Sub
    End If
    
    Set objPolicy = Nothing
    Set objDomain = Nothing

    Print "Created policy " & strPolicy

End Sub

'********************************************************************
'*
'* Sub DELETEDAP()
'* Purpose: Deletes a LDAP policy object.
'* Input:   strPolicy           the name of the policy object
'*          strCurrentUser      the name or cn of the current user
'*          strPassword         the current user password
'*          blnQuiet            specifies whether to suppress messages
'* Output:  None
'*
'********************************************************************

Sub DELETELDAP (strPolicy, strCurrentUser, strPassword, blnQuiet)
    ON ERROR RESUME NEXT

    Dim  objDomain, objProvider, strConfig, objPolicy 


    Set objDomain = GetObject("LDAP://RootDSE")
    strConfig = objDomain.Get("configurationNamingContext")
    Set objDomain=Nothing
   
    If Err.Number then
	Print "Error 0x" & CStr(Hex(Err.Number)) & " occurred in getting configuration container."
	If Err.Description <> "" Then Print "Error description: " & Err.Description & "."
	Err.Clear
        Exit Sub
    End If

    If strCurrentUser = "" Then            'no user credential is passed
        Set objDomain = GetObject("LDAP://CN=Query-Policies,CN=Directory Service,CN=Windows NT,CN=Services," & strConfig)
    Else
        Set objProvider = GetObject("LDAP:")
		'Use user authentication
        Set objDomain = objProvider.OpenDsObject("LDAP://CN=Query-Policies,CN=Directory Service,CN=Windows NT,CN=Services," _
                           & strConfig,strCurrentUser,strPassword,1)
    End If
    If Err.Number then
	Print "Error 0x" & CStr(Hex(Err.Number)) & " occurred in getting query policy container."
	If Err.Description <> "" Then Print "Error description: " & Err.Description & "."
	Err.Clear
        Exit Sub
    End If


    Set objPolicy = objDomain.GetObject("queryPolicy","cn=" &  strPolicy)
    If Err.Number then
	Print "Error 0x" & CStr(Hex(Err.Number)) & " occurred in getting object " & strPolicy & "."
	If Err.Description <> "" Then Print "Error description: " & Err.Description & "."
	Err.Clear
        Exit Sub
    End If

    if Not blnCheckForQueryPolicy(objPolicy.distinguishedName, strConfig, strCurrentUser, strPassword) Then
    	Print "Domain controllers or Sites are referencing the policy object " & strPolicy 
        Print "Cannot delete policies if they are in use by any domain controllers or sites"
        Exit Sub
    End If


    objDomain.Delete "queryPolicy","cn=" &  strPolicy
    If Err.Number then
	Print "Error 0x" & CStr(Hex(Err.Number)) & " occurred deleting object "	& strPolicy & "."
	If Err.Description <> "" Then Print "Error description: " & Err.Description & "."
	Err.Clear
        Exit Sub
    End If
    
    Set objPolicy = Nothing    
    Set objDomain = Nothing
    
    Print "Policy " & strPolicy & " deleted" 

End Sub

'********************************************************************
'*
'* Sub REMOVELDAP()
'* Purpose: Removes a query policy reference from a site or server.
'* Input:   strServer           the name of the site or server
'*          strCurrentUser      the name or cn of the current user
'*          strPassword         the current user password
'*          blnQuiet            specifies whether to suppress messages
'* Output:  None
'*
'********************************************************************

Sub REMOVELDAP (strServer, strCurrentUser, strPassword, blnQuiet)

    ON ERROR RESUME NEXT

    Dim  objDomain, objProvider, objServer, strConfig, objPolicy, strClass
    Dim  objNTDSSettings, strCriteria, strADSPath , i, blnSite

    blnSite = False

    Set objDomain = GetObject("LDAP://RootDSE")
    strConfig = objDomain.Get("configurationNamingContext")
    Set objDomain=Nothing
   
    If Err.Number then
	Print "Error 0x" & CStr(Hex(Err.Number)) & " occurred in getting configuration container."
	If Err.Description <> "" Then Print "Error description: " & Err.Description & "."
	Err.Clear
        Exit Sub
    End If

    strCriteria = "(&(|(objectClass=site)(objectClass=server))(cn=" & strServer & "))"
    
    
    if blnSearchForServer(strServer, strConfig, strCurrentUser, strPassword, strCriteria, strADSPath, strClass) Then

    For i = LBound(strClass) to UBound(strClass)
      If InStr(1, strClass(i), "site",1) then 
        blnSite = True
        Exit For
      End If  
    Next
    
    	If strCurrentUser = "" Then            'no user credential is passed
        	Set objServer = GetObject(strADSPath)
	Else
        	Set objProvider = GetObject("LDAP:")
		'Use user authentication
	        Set objServer = objProvider.OpenDsObject(strADSPath,strCurrentUser,strPassword,1)
	End If
    	If Err.Number then
		Print "Error 0x" & CStr(Hex(Err.Number)) & " occurred in getting object " & strADSPath & "."
		If Err.Description <> "" Then Print "Error description: " & Err.Description & "."
		Err.Clear
	        Exit Sub
    	End If
        
            
        If blnSite Then
        	Set objNTDSSettings = objServer.GetObject("nTDSSiteSettings", "cn=NTDS Site Settings")
        Else
                Set objNTDSSettings = objServer.GetObject("nTDSDSA", "cn=NTDS Settings")
        End If
        If Err.Number then
		Print "Error 0x" & CStr(Hex(Err.Number)) & " occurred in getting NTDS object " & strADSPath & "."
		If Err.Description <> "" Then Print "Error description: " & Err.Description & "."
		Err.Clear
	        Exit Sub
    	End If

        objNTDSSettings.PutEx ADS_PROPERTY_CLEAR, "queryPolicyObject",""
        objNTDSSettings.SetInfo

        If Err.Number then
		Print "Error 0x" & CStr(Hex(Err.Number)) & " occurred in setting NTDSDSDA object."
		If Err.Description <> "" Then Print "Error description: " & Err.Description & "."
		Err.Clear
                Print "Server " & strServer & " was not reset"
	        Exit Sub
        Else
           Print "Server " & strServer & " was reset"
    	End If
    Else
       Print "Server " & strServer & " was not found"
    End If

    Set objNTDSSettings = Nothing
    Set objServer = Nothing
    Set objDomain = Nothing
    
End Sub




'********************************************************************
'* Function blnCheckForQueryPolicy()
'* 
'* Purpose: query domain controllers to see if they are using the policy
'* Input: 	adspath 	path of the policy
'*		strSearchpath	base search should be config nc
'*		strCurrentUser	username
'*		strPassword	password
'*  blnCheckForQueryPolicy returns True if servers are not using policy
'*
'********************************************************************

Function blnCheckForQueryPolicy(ADSPath, strSearchPath, strCurrentUser, strPassword)

ON ERROR RESUME NEXT

Dim objConnect, objCommand, objRecordSet, intResult
Dim strProperties, strScope, strCriteria, strCommand


blnCheckForQueryPolicy = False
strScope = "SubTree"
strProperties = "cn"
strCriteria = "(&(|(objectClass=nTDSDSA)(objectClass=nTDSSiteSettings))(queryPolicyObject=" & ADSPath & "))"

    Set objConnect = CreateObject("ADODB.Connection")
    If Err.Number then
        Print "Error 0x" & CStr(Hex(Err.Number)) & " ocurred in opening a connection."
        If Err.Description <> "" Then Print "Error description: " & Err.Description & "."
        Exit Function
    End If

    Set objCommand = CreateObject("ADODB.Command")
    If Err.Number then
        Print "Error 0x" & CStr(Hex(Err.Number)) & " ocurred in creating the command object."
        If Err.Description <> "" Then Print "Error description: " & Err.Description & "."
        Exit Function
    End If

    objConnect.Provider = "ADsDSOObject"
    If strCurrentUser = "" then
        objConnect.Open "Active Directory Provider"
    Else
        objConnect.Open "Active Directory Provider", strCurrentUser, strPassword
    End If
    If Err.Number then
        Print "Error 0x" & CStr(Hex(Err.Number)) & " ocurred opening a provider."
        If Err.Description <> "" Then Print "Error description: " & Err.Description & "."
        Exit Function
    End If

    Set objCommand.ActiveConnection = objConnect

    'Set the query string and other properties
    strCommand = "<LDAP://" & strSearchPath & ">;" & strCriteria & ";" & strProperties & ";" & strScope
    objCommand.CommandText  = strCommand
    objCommand.Properties("Page Size") = 100000                    
    objCommand.Properties("Timeout") = 300000 'seconds


   Set objRecordSet = objCommand.Execute
   If Err.Number then
        Print "Error 0x" & CStr(Hex(Err.Number)) & " ocurred during the search."
        If Err.Description <> "" Then Print "Error description: " & Err.Description & "."
        Err.Clear
        Exit Function
   End If

   If objRecordSet.RecordCount = 0 Then blnCheckForQueryPolicy = True

End Function


'********************************************************************
'* Function blnSearchForServer()
'* 
'* Purpose: query config nc for requested domain controller
'* Input: 	strServer 	server name
'*		strSearchpath	base search should be config nc
'*		strCurrentUser	username
'*		strPassword	password
'* Output:      strADSPath      ADS Path to the domain controller
'*              strClass        Class of object
'*  blnCheckForQUeryPolicy returns True if servers are not using policy
'*
'********************************************************************

Function blnSearchForServer(strServer, strSearchPath, strCurrentUser, strPassword, strCriteria, strADSPath, strClass)

ON ERROR RESUME NEXT

Dim objConnect, objCommand, objRecordSet, intResult
Dim strProperties, strScope, strCommand


blnSearchForServer = False
strScope = "SubTree"
strProperties = "ADSPath,objectClass"


    Set objConnect = CreateObject("ADODB.Connection")
    If Err.Number then
        Print "Error 0x" & CStr(Hex(Err.Number)) & " ocurred in opening a connection."
        If Err.Description <> "" Then Print "Error description: " & Err.Description & "."
        Exit Function
    End If

    Set objCommand = CreateObject("ADODB.Command")
    If Err.Number then
        Print "Error 0x" & CStr(Hex(Err.Number)) & " ocurred in creating the command object."
        If Err.Description <> "" Then Print "Error description: " & Err.Description & "."
        Exit Function
    End If

    objConnect.Provider = "ADsDSOObject"
    If strCurrentUser = "" then
        objConnect.Open "Active Directory Provider"
    Else
        objConnect.Open "Active Directory Provider", strCurrentUser, strPassword
    End If
    If Err.Number then
        Print "Error 0x" & CStr(Hex(Err.Number)) & " ocurred opening a provider."
        If Err.Description <> "" Then Print "Error description: " & Err.Description & "."
        Exit Function
    End If

    Set objCommand.ActiveConnection = objConnect

    'Set the query string and other properties
    strCommand = "<LDAP://" & strSearchPath & ">;" & strCriteria & ";" & strProperties & ";" & strScope
    objCommand.CommandText  = strCommand
    objCommand.Properties("Page Size") = 100000                    
    objCommand.Properties("Timeout") = 300000 'seconds


   Set objRecordSet = objCommand.Execute
   If Err.Number then
        Print "Error 0x" & CStr(Hex(Err.Number)) & " ocurred during the search."
        If Err.Description <> "" Then Print "Error description: " & Err.Description & "."
        Err.Clear
        Exit Function
    End If

    If objRecordSet.RecordCount = 1 Then 
	blnSearchForServer = True
        strADSPath = objRecordSet.Fields(0).Value
        strClass = objRecordSet.Fields(1).Value
    End If

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
        Print "Argument is not an array!"
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
'* Purpose:   Prints a message on screen if blnQuiet = False.
'* Input:     strMessage    the string to print
'* Output:    strMessage is printed on screen if blnQuiet = False.
'*
'********************************************************************

Sub Print(ByRef strMessage)
    'If Not blnQuiet then
        Wscript.Echo  strMessage
    'End If
End Sub

'********************************************************************
'*                                                                  *
'*                           End of File                            *
'*                                                                  *
'********************************************************************

'********************************************************************
'*
'* Procedures calling sequence: MODIFYLDAP.VBS
'*
'*  intChkProgram
'*	intParseCmdLine
'*	ShowUsage
'*	MODIFYLDAP
'*      PRINTLDAP
'*      ASSIGNLDAP
'*      	blnSearchForServer
'*      DELETELDAP
'*      	blnCheckForQueryPolicy
'*      CREATELDAP
'*      REMOVELDAP
'*
'********************************************************************
