set Wshshell = WScript.CreateObject("WScript.Shell")

WshShell.RegWrite "HKLM\Software\Microsoft\PCHealth\ErrorReporting\AllOrNone", "0", "REG_DWORD"

WshShell.RegWrite "HKLM\Software\Microsoft\PCHealth\ErrorReporting\IncludeMicrosoftApps", "0", "REG_DWORD"

WshShell.RegWrite "HKLM\Software\Microsoft\PCHealth\ErrorReporting\IncludeWindowsApps", "0", "REG_DWORD"

WshShell.RegWrite "HKLM\Software\Microsoft\PCHealth\ErrorReporting\IncludeKernelFaults", "0", "REG_DWORD"

WshShell.RegWrite "HKLM\Software\Microsoft\PCHealth\ErrorReporting\DoReport", "0", "REG_DWORD"





