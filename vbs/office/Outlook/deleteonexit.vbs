set Wshshell = WScript.CreateObject("WScript.Shell")

WshShell.RegWrite "HKCU\Software\Microsoft\Office\10.0\Outlook\Options\General\Warndelete", "1", "REG_DWORD"