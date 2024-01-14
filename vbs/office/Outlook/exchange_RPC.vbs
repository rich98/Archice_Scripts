set Wshshell = WScript.CreateObject("WScript.Shell")

WshShell.RegWrite "HKLM\Software\Microsoft\Exchange\Exchange Provider\Rpc_Binding_order", "ncalrpc,ncacn_ip_tcp,ncacn_spx,netbios,ncacn_vns_spp", "REG_SZ"

WshShell.RegWrite "HKLM\Software\Microsoft\Rpc\ClientProtocols\ncacn_http", "Rpcrt4.dll", "REG_SZ"

WshShell.RegWrite "HKLM\Software\Microsoft\Rpc\ClientProtocols\ncacn_ip_tcp", "Rpcrt4.dll", "REG_SZ"

WshShell.RegWrite "HKLM\Software\Microsoft\Rpc\ClientProtocols\ncacn_np", "Rpcrt4.dll", "REG_SZ"

WshShell.RegWrite "HKLM\Software\Microsoft\Rpc\ClientProtocols\ncacn_nb_tcp", "Rpcrt4.dll", "REG_SZ"

WshShell.RegWrite "HKLM\Software\Microsoft\Rpc\ClientProtocols\ncacn_ip_udp", "Rpcrt4.dll", "REG_SZ"



