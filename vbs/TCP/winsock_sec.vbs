' VBScript source code
' Created with Visual Studio.Net
' Ben Smith - Microsoft Corporation
' Microsoft Windows Security Resource Kit
' Registry script - Winsock Security Options for Windows 2000/Windows XP
' Version 1.1

'require variable declaration

option explicit

' declare varaibles

dim oShell

' main

set oShell = createobject("Wscript.shell")

'set winsock security registry entries

oShell.RegWrite "HKLM\System\CurrentControlSet\Services\AFD\Parameters\EnableDynamicBacklog",1,"REG_DWORD"
oShell.RegWrite "HKLM\System\CurrentControlSet\Services\AFD\Parameters\DynamicBacklogGrowthDelta",10,"REG_DWORD"
oShell.RegWrite "HKLM\System\CurrentControlSet\Services\AFD\Parameters\MinimumDynamicBacklog",20,"REG_DWORD"
oShell.RegWrite "HKLM\System\CurrentControlSet\Services\AFD\Parameters\MaximumDynamicBacklog",20000,"REG_DWORD"


wscript.echo ("Winsock Security Options Set")

set oShell = nothing

