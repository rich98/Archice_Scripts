set Wshshell = WScript.CreateObject("WScript.Shell")

' WshShell.RegWrite "HKLM\Software\Microsoft\Internet Explorer\Control Panel\", "Control Panel"


' WshShell.RegWrite "HKLM\Software\Microsoft\Internet Explorer\Control Panel\AdvancedTab", "8080", "REG_DWORD"

WshShell.Regwrite "HKLM\SOFTWARE\Network Associates\TVD\Shared Components\McUpdate\CurrentVersion\Update\Update Site1\dwProxyPort", "8080", "REG_DWORD"

'WshShell.Regwrite "HKLM\SOFTWARE\Network Associates\TVD\Shared Components\McUpdate\CurrentVersion\Update\Update Site1\dwProxyPort", ""


'WshShell.Regwrite "HKLM\SOFTWARE\Network Associates\TVD\Shared Components\McUpdate\CurrentVersion\Update\Update Site1\szProxy", ""

WshShell.Regwrite "HKLM\SOFTWARE\Network Associates\TVD\Shared Components\McUpdate\CurrentVersion\Update\Update Site1\szProxy", "VPN1"