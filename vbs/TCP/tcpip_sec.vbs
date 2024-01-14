' VBScript source code
' Created with Visual Studio.Net
' Ben Smith - Microsoft Corporation
' Microsoft Windows Security Resource Kit
' Registry script - TCP/IP Security Options for Windows 2000/Windows XP
' Version 1.1

'require variable declaration

option explicit

' declare variables

dim oShell

' main

set oShell = createobject("Wscript.shell")

'set TCP/IP security registry entries

oShell.RegWrite "HKLM\System\CurrentControlSet\Services\Tcpip\Parameters\EnableICMPRedirect",0,"REG_DWORD"
oShell.RegWrite "HKLM\System\CurrentControlSet\Services\Tcpip\Parameters\SynAttackProtect",2,"REG_DWORD"
oShell.RegWrite "HKLM\System\CurrentControlSet\Services\Tcpip\Parameters\TcpMaxConnectResponseRetransmissions",2,"REG_DWORD"
oShell.RegWrite "HKLM\System\CurrentControlSet\Services\Tcpip\Parameters\TCPMaxHalfOpen",500,"REG_DWORD"
oShell.RegWrite "HKLM\System\CurrentControlSet\Services\Tcpip\Parameters\TCPMaxHalfOpenRetired",400,"REG_DWORD"
oShell.RegWrite "HKLM\System\CurrentControlSet\Services\Tcpip\Parameters\TCPMaxPortsExhausted",5,"REG_DWORD"
oShell.RegWrite "HKLM\System\CurrentControlSet\Services\Tcpip\Parameters\TcpMaxDataRetransmissions",3,"REG_DWORD"
oShell.RegWrite "HKLM\System\CurrentControlSet\Services\Tcpip\Parameters\EnableDeadGWDetect",0,"REG_DWORD"
oShell.RegWrite "HKLM\System\CurrentControlSet\Services\Tcpip\Parameters\EnablePMTUDiscovery",0,"REG_DWORD"
oShell.RegWrite "HKLM\System\CurrentControlSet\Services\Tcpip\Parameters\KeepAliveTime",300000,"REG_DWORD"
oShell.RegWrite "HKLM\System\CurrentControlSet\Services\Tcpip\Parameters\DisableIPSourceRouting",2,"REG_DWORD"
oShell.RegWrite "HKLM\System\CurrentControlSet\Services\Tcpip\Parameters\NoNameReleaseOnDemand",1,"REG_DWORD"
oShell.RegWrite "HKLM\System\CurrentControlSet\Services\Tcpip\Parameters\PerformRouterDiscovery",0,"REG_DWORD"

wscript.echo ("TCP/IP Security Options Set")

set oShell = nothing

