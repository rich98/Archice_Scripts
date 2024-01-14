set Wshshell = WScript.CreateObject("WScript.Shell")

WshShell.RegWrite "HKCU\Software\Policies\Microsoft\Security\CheckAdminSettings", "00000001"

