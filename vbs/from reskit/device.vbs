'********************************************************************
'*
'* File:           Device.vbs
'* Created:        March 1999
'* Version:        1.0
'*
'*  Main Function:  Controls Devices on a machine.
'*
'*    1.  Device.vbs /L"
'*                   [/S <server>][/U <username>][/W <password>]"
'*                   [/O <outputfile>]"
'*    
'*    2.  Device.vbs /G | /X | /R | /M <StartMode>"
'*                   /D <device>"
'*                   [/S <server>][/U <username>][/W <password>]"
'*                   [/O <outputfile>]"
'*    
'*    3.  Device.vbs /I /D <device> [/N <DisplayName>]"
'*                   [/S <server>][/U <username>][/W <password>]"
'*                   [/O <outputfile>]"
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
    CONST CONST_LIST                    = "LIST"
    CONST CONST_STOP                    = "STOP"
    CONST CONST_START                   = "START"
    CONST CONST_MODE                    = "MODE"
    CONST CONST_INSTALL                 = "INSTALL"
    CONST CONST_REMOVE                  = "REMOVE"
    CONST CONST_DEFAULTTASK             = "DEVICE"


    'Declare variables
    Dim intOpMode,   i
    Dim strServer, strUserName, strPassword, strOutputFile
    Dim strTaskCommand,   strDriverName,    strStartMode,      strDisplayName
    Dim blnDetails

    'Make sure the host is csript, if not then abort
    VerifyHostIsCscript()

    'Parse the command line
    intOpMode = intParseCmdLine(strServer      ,  _
                                strUserName    ,  _
                                strPassword    ,  _
                                strOutputFile  ,  _
                                strTaskCommand ,  _
                                strDriverName  ,  _
                                strStartMode   ,  _
                                strDisplayName ,  _
                                blnDetails        )

    Select Case intOpMode

        Case CONST_SHOW_USAGE
            Call ShowUsage()

        Case CONST_PROCEED
            Call DEVICE(strServer      ,  _
                        strUserName    ,  _
                        strPassword    ,  _
                        strOutputFile  ,  _
                        strTaskCommand ,  _ 
                        strDriverName  ,  _
                        strStartMode   ,  _
                        strDisplayName ,  _
                        blnDetails        )

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
'* Sub DEVICE()
'* Purpose: Controls DEVICEs on a machine.
'* Input:   
'*          strServer          a machine name
'*          strOutputFile      an output file name
'*          strUserName        the current user's name
'*          strPassword        the current user's password
'*          strTaskCommand     one of /list, /start, /stop /install /remove
'*                                    /dependents
'*          strDriverName      name of the DEVICE
'*          strStartMode       start mode of the DEVICE
'*          strDisplayName     Display name for the DEVICE.
'*          blnDetails         Extra information to be displayed on the output
'* Output:  Results are either printed on screen or saved in strOutputFile.
'*
'********************************************************************

Private Sub DEVICE(strServer      ,  _
                   strUserName    ,  _
                   strPassword    ,  _
                   strOutputFile  ,  _
                   strTaskCommand ,  _ 
                   strDriverName  ,  _
                   strStartMode   ,  _
                   strDisplayName ,  _
                   blnDetails        )

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

        Case CONST_LIST
            Call DeviceList(objService, objOutputFile, blnDetails)
        Case CONST_START
            Call DeviceStart(objService, objOutputFile, strDriverName)
        Case CONST_STOP
            Call DeviceStop(objService, objOutputFile, strDriverName)
        Case CONST_MODE
            Call DeviceMode(objService, objOutputFile, strDriverName, _
                            strStartMode)
        Case CONST_INSTALL
            Call DeviceInstall(objService, objOutputFile, strDriverName, _
                               strDisplayName)
        Case CONST_REMOVE
            Call DeviceRemove(objService, objOutputFile, strDriverName)

    End Select

    If NOT IsEmpty(objOutputFile) Then
        objOutputFile.Close
        Wscript.Echo "Results are saved in file " & strOutputFile & "."
    End If


End Sub

'********************************************************************
'*
'* Sub DeviceStart()
'* Purpose: Starts a driver.
'* Input:   objService         a Device object
'*          objOutputFile      an output file object
'*          strDriverName      name of the Device to be started or stopped
'* Output:  Results are either printed on screen or saved in objOutputFile.
'*
'********************************************************************
Private Sub DeviceStart(objService, objOutputFile, strDriverName)

    ON ERROR RESUME NEXT

    Dim objEnumerator, objInstance
    Dim strMessage
    Dim intStatus

    strMessage        = ""
    
    Set objInstance = objService.Get("Win32_SystemDriver='" &_
                                      strDriverName & "'")
    If Err.Number Then
        If err.number = -2147217406 then ' Invalid Device Name
		    call Print("Device " & strDriverName & " is not valid.")
		    call Print("Check for valid Device names with the " _
                             & "/LIST switch.")
		Else
            call Print( "Error 0x" & CStr(Hex(Err.Number)) & " occurred in " _
                      & "getting device " & strDriverName & ".")	
            If Err.Description <> "" Then
                 call Print( "Error description: " & Err.Description & ".")
            End If
        End If
        Err.Clear
        Exit Sub
    End If
    If objInstance is nothing Then
        Exit Sub
    Else
        intStatus = objInstance.StartService()
     	if blnErrorOccurred("Provider Failure.") Then
		    call Print("Device " & strDriverName & " is not valid.")
         	Exit Sub
        End If                
        If intStatus = 0 Then
            strMessage = "Succeeded in starting device " & strDriverName & "."
        Else
            strMessage = "Failed to start device " & strDriverName & "."
        End If
        WriteLine strMessage, objOutputFile
    End If

End Sub

'********************************************************************
'*
'* Sub DeviceStop()
'* Purpose: Stops a Driver.
'* Input:   objService          a Device object
'* Purpose: Starts a driver.
'* Input:   objService         a Device object
'*          objOutputFile      an output file object
'*          strDriverName      name of the Device to be started or stopped
'* Output:  Results are either printed on screen or saved in objOutputFile.
'*
'********************************************************************
Private Sub DeviceStop(objService, objOutputFile, strDriverName)

    ON ERROR RESUME NEXT

    Dim objEnumerator, objInstance
    Dim strMessage
    Dim intStatus
    
    strMessage        = ""
    
    Set objInstance = objService.Get("Win32_SystemDriver='" _
                                     & strDriverName & "'")
    If Err.Number Then
        If err.number = -2147217406 then ' Invalid Device Name
		    call Print("Device " & strDriverName & " is not valid.")
		    call Print("Check for valid Device names with the " _
                             & "/LIST switch.")
		Else
            call Print( "Error 0x" & CStr(Hex(Err.Number)) & " occurred in " _
                      & "getting device " & strDriverName & ".")	
            If Err.Description <> "" Then
                 call Print( "Error description: " & Err.Description & ".")
            End If
        End If
        Err.Clear
        Exit Sub
    End If
    If objInstance is nothing Then
        Exit Sub
    Else
        intStatus = objInstance.StopService()
     	if blnErrorOccurred("Provider Failure.") Then
            call Print("Check for valid Device names with the /LIST switch.")
            Exit Sub
        End If
        If intStatus = 0 Then
            strMessage = "Succeeded in stopping Driver " & strDriverName & "."
        Else
            strMessage = "Failed to stop Driver " & strDriverName & "."
        End If
        WriteLine strMessage, objOutputFile
    End If

End Sub

'********************************************************************
'*
'* Sub DeviceMode()
'* Purpose: Sets the startup mode of a device.
'* Input:   objService          a Device object
'*          objOutputFile       an output file object
'*          strDriverName       name of the Device to be started or stopped
'*          strStartMode        The Mode to set the device to
'*
'* Output:  Results are either printed on screen or saved in objOutputFile.
'*
'********************************************************************
Private Sub DeviceMode(objService, objOutputFile, strDriverName, strStartMode)

    ON ERROR RESUME NEXT

    Dim objEnumerator, objInstance
    Dim strMessage
    Dim intStatus
    
    strMessage        = ""
    
    Set objInstance = objService.Get("Win32_SystemDriver='"& strDriverName&"'")
    If Err.Number Then
        If err.number = -2147217406 then ' Invalid Device Name
            call Print("Device " & strDriverName & " is not valid.")
            call Print("Check for valid Device names with the /LIST switch.")
		Else
            call Print( "Error 0x" & CStr(Hex(Err.Number)) & " occurred in " _
                      & "getting device " & strDriverName & ".")	
            If Err.Description <> "" Then
                 call Print( "Error description: " & Err.Description & ".")
            End If
        End If
        Err.Clear
        Exit Sub
    End If
    If objInstance is nothing Then
        Exit Sub
    Else
        intStatus = objInstance.ChangeStartMode(strStartMode)
     	if blnErrorOccurred("Provider Failure.") Then
            Call Print("Check for valid Device names with the /LIST switch.")
            Exit Sub
        End If
        If intStatus = 0 Then
            strMessage = "Succeeded in changing start mode of the Device " _
                          & strDriverName & "."
        Else
            strMessage = "Failed to change the start mode of the Device " _
                          & strDriverName & "."
        End If
        WriteLine strMessage, objOutputFile
    End If

End Sub

'********************************************************************
'*
'* Sub DeviceInstall()
'* Purpose: Installs a Driver.
'* Input:   objService          a Device object
'*          objOutputFile       an output file object
'*          strDriverName       name of the Device to be started or stopped
'*          strDisplayName      The Name displayed on the driver list
'*
'* Output:  Results are either printed on screen or saved in objOutputFile.
'*
'********************************************************************
Private Sub DeviceInstall(objService    ,  _
                          objOutputFile ,  _
                          strDriverName ,  _
                          strDisplayName   )

    ON ERROR RESUME NEXT

    Dim objEnumerator, objInstance
    Dim strMessage
    Dim intStatus
    
    strMessage        = ""
    
    Set objInstance = objService.Get("Win32_SystemDriver")
    If Err.Number Then
        call Print( "Error 0x" & CStr(Hex(Err.Number)) & " occurred in " _
                  & "getting Device " & strDriverName & ".")
        If Err.Description <> "" Then
            call Print( "Error description: " & Err.Description & ".")
        End If
        Err.Clear
        Exit Sub
    End If
    If objInstance is Nothing Then
        Exit Sub
    Else

        If IsEmpty(strDisplayName) then strDisplayName = strDriverName

        intStatus = objInstance.Create(strDriverName, strDisplayName, _
                                       strDriverName)
     	if blnErrorOccurred("Provider Failure.") Then
            call Print("Valid Driver name not specified.")
            Exit Sub
        End If
        If intStatus = 0 Then
            strMessage = "Succeeded in creating Device " & strDriverName & "."
        Else
            strMessage = "Failed to create Device " & strDriverName & "."
        End If
        WriteLine strMessage, objOutputFile
    End If

End Sub

'********************************************************************
'*
'* Sub DeviceRemove()
'* Purpose: Removes a Driver.
'* Input:   objService          a Device object
'*          objOutputFile       an output file object
'*          strDriverName       name of the Device to be started or stopped
'*
'* Output:  Results are either printed on screen or saved in objOutputFile.
'*
'********************************************************************
Private Sub DeviceRemove(objService, objOutputFile, strDriverName)

    ON ERROR RESUME NEXT

    Dim objEnumerator, objInstance
    Dim strMessage
    Dim intStatus
    
    strMessage        = ""
 
    Set objInstance = objService.Get("Win32_SystemDriver='" _
                                    & strDriverName&"'")
    If Err.Number Then
        If err.number = -2147217406 then ' Invalid Device Name
                call Print("Device " & strDriverName & " is not valid.")
                call Print("Check for valid Device names with the " _
                         & "/LIST switch.")
        Else
            call Print( "Error 0x" & CStr(Hex(Err.Number)) & " occurred in " _
                      & "getting device " & strDriverName & ".")	
            If Err.Description <> "" Then
                 call Print( "Error description: " & Err.Description & ".")
            End If
        End If
        Err.Clear
        Exit Sub
    End If
    If objInstance is Nothing Then
        Exit Sub
    Else
        intStatus = objInstance.Delete()
     	if blnErrorOccurred("Provider Failure.") Then
            call Print("Valid Driver name not specified.")
            Exit Sub
        End If        
        If intStatus = 0 Then
            strMessage = "Succeeded in deleting Device " & strDriverName & "."
        Else
            strMessage = "Failed to delete Device " & strDriverName & "."
        End If
        WriteLine strMessage, objOutputFile
    End If
            
End Sub

'********************************************************************
'*
'* Sub DeviceList()
'* Purpose: Lists all devices.
'* Input:   objService          a Device object
'*          objOutputFile       an output file object
'*          blnDetails          The option to display more information.
'*
'* Output:  Results are either printed on screen or saved in objOutputFile.
'*
'********************************************************************
Private Sub DeviceList(objService, objOutputFile, blnDetails)

    ON ERROR RESUME NEXT

    Dim objEnumerator, objInstance
    Dim strMessage
    Dim intDeviceNameLength
    ReDim strName(0), strDisplayName(0), strState(0), strStartModeDsp(0)
    ReDim intOrder(0)

	'Initialize local variables
    strMessage        = ""
    strName(0)        = ""
    strDisplayName(0) = ""
    strState(0)       = ""
    intOrder(0)       = 0

    Set objEnumerator = objService.ExecQuery ( _
                        "Select Name,PathName,State, StartMode From " _
                      & "Win32_SystemDriver",,0)
    If Err.Number Then
        call Print( "Error 0x" & CStr(Hex(Err.Number)) _
                  & " occurred during the query.")
        If Err.Description <> "" Then
            call Print( "Error description: " & Err.Description & ".")
        End If
        Err.Clear
        Exit Sub
    End If
    i = 0
    For Each objInstance in objEnumerator
        If objInstance is nothing Then
            Exit Sub
        Else
            ReDim Preserve strName(i), strDisplayName(i), _
                           strState(i), strStartModeDsp(i), intOrder(i)
            strName(i) = objInstance.Name
            strDisplayName(i) = objInstance.PathName
            strState(i) = objInstance.State
			strStartModeDsp(i) = objInstance.StartMode
            intOrder(i) = i
            i = i + 1
        End If
        If Err.Number Then
            Err.Clear
        End If
    Next
	intDeviceNameLength = 12
    If i > 0 Then
		'Check Spacing
        For i = 0 To UBound(strName)
            If len(strName(i)) > (intDeviceNameLength - 1) then
                Do until (intDeviceNameLength - 1) > len(strName(i))
                    intDeviceNameLength = intDeviceNameLength + 1
                Loop
            End If
        Next

        'Display the header
        strMessage = Space(2) & strPackString("NAME", intDeviceNameLength, 1, 1)
        strMessage = strMessage & strPackString("STATE", 10, 1, 1)
        strMessage = strMessage & strPackString("STARTUP", 10, 1, 0)
            If BlnDetails = True then
                strMessage = strMessage & strPackString("PATH NAME", 15, 1, 0) _
                             & vbCRLF
            End IF
        WriteLine strMessage, objOutputFile
        Call SortArray(strName, True, intOrder, 0)
        Call ReArrangeArray(strDisplayName, intOrder)
        Call ReArrangeArray(strState, intOrder)
		Call ReArrangeArray(strStartModeDsp, intOrder)
        For i = 0 To UBound(strName)
            strMessage = Space(2) & strPackString(strName(i), _
                         intDeviceNameLength, 1, 1)
            strMessage = strMessage & strPackString(strState(i), 10, 1, 1)
            strMessage = strMessage & strPackString(strStartModeDsp(i), _
                         10, 1, 0)
            If BlnDetails = True then
                strMessage = strMessage & strPackString(strDisplayName(i), _
                         15, 1, 0)
            End If
            WriteLine strMessage, objOutputFile
        Next
    Else
        Wscript.Echo "Device not found!"
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
                                  ByRef strDriverName  ,  _
                                  ByRef strStartMode   ,  _
                                  ByRef strDisplayName ,  _
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

            Case "/d"
                If Not blnGetArg ("driver name", strDriverName, intArgIter) Then
                    intParseCmdLine = CONST_ERROR
                    Exit Function
                End If
                intArgIter = intArgIter + 1
               
            Case "/g"
                intParseCmdLine = CONST_PROCEED
                strTaskCommand = CONST_START
                intArgIter = intArgIter + 1

            Case "/x"
                intParseCmdLine = CONST_PROCEED
                strTaskCommand = CONST_STOP
                intArgIter = intArgIter + 1
              
            Case "/m"
                intParseCmdLine = CONST_PROCEED
                strTaskCommand = CONST_MODE
                If Not blnGetArg ("start mode", strStartMode, intArgIter) Then
                    intParseCmdLine = CONST_ERROR
                    Exit Function
                End If
                intArgIter = intArgIter + 1

            Case "/i"
                intParseCmdLine = CONST_PROCEED
                strTaskCommand = CONST_INSTALL
                intArgIter = intArgIter + 1 

            Case "/r"
                intParseCmdLine = CONST_PROCEED
                strTaskCommand = CONST_REMOVE
                intArgIter = intArgIter + 1

            Case "/n"
                If blnGetArg ("display name", strDisplayName, intArgIter) Then
                    intParseCmdLine = CONST_ERROR
                    Exit Function
                End If
                intArgIter = intArgIter + 1
              
            Case "/l"
                intParseCmdLine = CONST_PROCEED
                strTaskCommand = CONST_LIST
                intArgIter = intArgIter + 1

            Case "/v"
                blnDetails = True
                intArgIter = intArgITer + 1

            Case Else 'We shouldn't get here
                Call Wscript.Echo("Invalid or misplaced parameter: " _
                   & Wscript.arguments.Item(intArgIter) & vbCRLF _
                   & "Please check the input and try again," & vbCRLF _
                   & "or invoke with '/?' for help with the syntax.")
                Wscript.Quit

        End Select

    Loop '** intArgIter <= Wscript.arguments.Count - 1

    If IsEmpty(intParseCmdLine) Then _
        intParseCmdLine = CONST_LIST

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
    Wscript.Echo "Controls Devices on a machine."
    Wscript.Echo ""
    Wscript.Echo "SYNTAX:"
    Wscript.Echo "1.  Device.vbs /L"
    Wscript.Echo "               [/S <server>][/U <username>][/W <password>]"
    Wscript.Echo "               [/O <outputfile>]"
    Wscript.Echo ""
    Wscript.Echo "2.  Device.vbs /G | /X | /R | /M <StartMode>"
    Wscript.Echo "               /D <device>"
    Wscript.Echo "               [/S <server>][/U <username>][/W <password>]"
    Wscript.Echo "               [/O <outputfile>]"
    Wscript.Echo ""
    Wscript.Echo "3.  Device.vbs /I /D <device> [/N <DisplayName>]"
    Wscript.Echo "               [/S <server>][/U <username>][/W <password>]"
    Wscript.Echo "               [/O <outputfile>]"
    Wscript.Echo ""
    Wscript.Echo "PARAMETER SPECIFIERS:"
    Wscript.Echo "   /L            List all devices"
    Wscript.Echo "   /G            Start a device"
    Wscript.Echo "   /X            Stop a device"
    Wscript.Echo "   /R            Remove a device"
    Wscript.Echo "   /M            Set the device Mode"
    Wscript.Echo "   /I            Install device"
    Wscript.Echo "   StartMode     The Device Startup Setting."
    Wscript.Echo "   device        The Full name and path of the device."
    Wscript.Echo "   DisplayName   The Device name that appears in the" _
               & " directory listing/"
    Wscript.Echo "   server        A machine name."
    Wscript.Echo "   username      The current user's name."
    Wscript.Echo "   password      Password of the current user."
    Wscript.Echo "   outputfile    The output file name."
    Wscript.Echo ""
    Wscript.Echo "EXAMPLE:"
    Wscript.Echo "1. cscript Device.vbs /L /S MyMachine2"
    Wscript.Echo "   Listed installed devices for the machine MyMachine2."
    Wscript.Echo "2. cscript Device.vbs /X /D Beep"
    Wscript.Echo "   Stops device Beep on the current machine."
    Wscript.Echo ""

End Sub

'********************************************************************
'* General Routines
'********************************************************************

'********************************************************************
'*
'* Sub SortArray()
'* Purpose: Sorts an array and arrange another array accordingly.
'* Input:   strArray    the array to be sorted
'*          blnOrder    True for ascending and False for descending
'*          strArray2   an array that has exactly the same number of 
'*                      elements as strArray and will be reordered 
'*                      together with strArray
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
        call Print( "Argument is not an array!")
        Exit Sub
    End If

    blnOrder = CBool(blnOrder)
    blnCase = CBool(blnCase)
    If Err.Number Then
        call Print( "Argument is not a boolean!")
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
            Else                 'Not case sensitive
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
        call Print( "At least one of the arguments is not an array")
        Exit Sub
    End If

    intUBound = UBound(strArray)

    If intUBound <> UBound(intOrder) Then
        call Print( "The upper bound of these two arrays do not match!")
        Exit Sub
    End If

    ReDim strTempArray(intUBound)

    For i = 0 To intUBound
        strTempArray(i) = strArray(intOrder(i))
        If Err.Number Then
            call Print( "Error 0x" & CStr(Hex(Err.Number)) & " occurred in " _
                      & "rearranging an array.")
            If Err.Description <> "" Then
                call Print( "Error description: " & Err.Description & ".")
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
             "1. Using ""CScript Device.vbs arguments"" for Windows 95/98 or" _
             & vbCRLF & "2. Changing the default Windows Scripting Host " _
             & "setting to CScript" & vbCRLF & "    using ""CScript " _
             & "//H:CScript //S"" and running the script using" & vbCRLF & _
             "    ""Device.vbs arguments"" for Windows NT/2000." )
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