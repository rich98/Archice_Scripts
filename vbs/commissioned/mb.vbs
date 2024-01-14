set Wshshell = WScript.CreateObject("WScript.Shell")

WshShell.RegWrite "HKLM\System\ControlSet001\Services\Browser\Parameters\MaintainServerList", "no"