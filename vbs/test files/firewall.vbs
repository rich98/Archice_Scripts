OPTION EXPLICIT

DIM ICSSC_DEFAULT, CONNECTION_PUBLIC, CONNECTION_PRIVATE, CONNECTION_ALL
DIM NetSharingManager
DIM PublicConnection, PrivateConnection
DIM EveryConnectionCollection

DIM objArgs
DIM con 

ICSSC_DEFAULT         = 0
CONNECTION_PUBLIC     = 0
CONNECTION_PRIVATE    = 1
CONNECTION_ALL        = 2

Main( )

sub Main( )
    Set objArgs = WScript.Arguments

    if objArgs.Count = 1 then
        con = objArgs(0)
        
        WScript.Echo con

        if Initialize() = TRUE then    
            GetConnectionObjects()

            FirewallTestByName(con)
        end if
    else
        DIM szMsg
        szMsg = "Invalid usage! Please provide the name of the connection as the argument." & chr(13) & chr(13) & _
                "Usage:" & chr(13) & _ 
                "       " + WScript.scriptname + " " + chr(34) + "Connection Name" + chr(34)
        WScript.Echo( szMsg )                
    end if

end sub


sub FirewallTestByName(conName)
on error resume next
    DIM Item
    DIM EveryConnection
    DIM objNCProps
    DIM szMsg
    DIM bFound
    
    bFound = false        
    for each Item in EveryConnectionCollection
        set EveryConnection = NetSharingManager.INetSharingConfigurationForINetConnection(Item)
        set objNCProps = NetSharingManager.NetConnectionProps(Item)
        if (ucase(conName) = ucase(objNCProps.Name)) then
            szMsg = "Enabling Firwall on connection:" & chr(13) & _
                    "Name: "       & objNCProps.Name & chr(13) & _
                    "Guid: "       & objNCProps.Guid & chr(13) & _
                    "DeviceName: " & objNCProps.DeviceName & chr(13) & _
                    "Status: "     & objNCProps.Status & chr(13) & _
                    "MediaType: "  & objNCProps.MediaType
            
            WScript.Echo(szMsg)
            bFound = true
            EveryConnection.EnableInternetFirewall
            exit for
        end if
    next
    
    if( bFound = false ) then
        WScript.Echo( "Connection " & chr(34) & conName & chr(34) & " was not found" )
    end if
    
end sub
