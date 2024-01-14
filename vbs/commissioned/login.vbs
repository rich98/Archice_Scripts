' login script for Forum Users

' login script order

' Network Drives
' Colour printers
' Proxy and port for av
' Start apps faste, fsater start menu, Clear payfile ion dhutdown
' audit
' outlook security fix for forms
' delete on exit (outlook 2002)
' balloon tips off
' Master Browser off
' NTFS timp stamp XP clients only
' Unload DLL's that stay in mem when not needed
' internet explorer cache size 1MB


' ***** Network drives *****

Set WshNetwork = CreateObject("WScript.Network")

On Error Resume Next

WshNetwork.RemoveNetworkDrive "t:" ,boolForce

WshNetwork.MapNetworkDrive "t:","\\forum10\share"

' ***** set printer path for colour printers *****

PrinterPath = "\\forum-print\GENICOM COLOUR cl160"

WshNetwork.AddWindowsPrinterConnection PrinterPath

PrinterPath = "\\forum-print\HP 8500 Colour - PS"

WshNetwork.AddWindowsPrinterConnection PrinterPath

WScript.Sleep 3000

' ***** set proxy and port values for av *****

Set WshShell = WScript.CreateObject ("WScript.Shell")

WshShell.Regwrite "HKLM\SOFTWARE\Network Associates\TVD\Shared Components\McUpdate\CurrentVersion\Update\Update Site1\dwProxyPort", "8080", "REG_DWORD"

WshShell.Regwrite "HKLM\SOFTWARE\Network Associates\TVD\Shared Components\McUpdate\CurrentVersion\Update\Update Site1\szProxy", "proxy"

' ***** Start apps faster and faster start menu*****

WshShell.RegWrite "HKLM\SYSTEM\CurrentControlSet\Control\Session Manager\Memory Management\PrefetchParameters\EnablePrefetcher", "30", "REG_DWORD"

WshShell.Regwrite "HKLM\SYSTEM\CurrentControlSet\Control\Session Manager\Memory Management\ClearPageFileAtShutdown", "1", "REG_DWORD"


WshShell.Regwrite "HKCU\Control Panel\Desktop\MenuShowDelay", "100"

' ***** Remove proxy and port values *****

'WshShell.Regwrite "HKLM\SOFTWARE\Network Associates\TVD\Shared Components\McUpdate\CurrentVersion\Update\Update Site1\dwProxyPort", ""

'WshShell.Regwrite "HKLM\SOFTWARE\Network Associates\TVD\Shared Components\McUpdate\CurrentVersion\Update\Update Site1\szProxy", "" 

' ***** Audit *****

' Set WshShell = WScript.CreateObject ("WScript.Shell")

' WshShell.Run "t:\it\track-it 6\audit32.exe"

set Wshshell = WScript.CreateObject("WScript.Shell")

' ***** Outlook security for forms *****

' WshShell.RegWrite "HKCU\Software\Policies\Microsoft\Security\CheckAdminSettings", "0000001"

' ***** Delete on exit *****

WshShell.RegWrite "HKCU\Software\Microsoft\Office\10.0\Outlook\Preferences\EmptyTrash", "1", "REG_DWORD"

WshShell.RegWrite "HKCU\Software\Microsoft\Office\10.0\Outlook\Options\General\Warndelete", "1", "REG_DWORD"

' ***** balloon tips off *****

set Wshshell = WScript.CreateObject("WScript.Shell")


WshShell.RegWrite "HKCU\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced\", "EnableBalloontips"


WshShell.RegWrite "HKCU\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced\EnableBalloontips", "0", "REG_DWORD"

' ***** Master browser off *****

WshShell.RegWrite "HKLM\System\ControlSet001\Services\Browser\Parameters\MaintainServerList", "no"

' ***** NTFS Time Stamp *****

WshShell.RegWrite "HKLM\SYSTEM\CurrentControlSet\Control\FileSystem\NtfsDisableLastAccessUpdate", "1", "REG_DWORD" 

' ***** unload DLL *****

WshShell.RegWrite "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\AlwaysUnloadDLL", "1", "REG_DWORD" 

' ***** IE cache size *****

WshShell.RegWrite "HKCU\Software\Microsoft\Windows\CurrentVersion\Internet Settings\5.0\Cache\Content\CacheLimit", "1024", "REG_DWORD"







