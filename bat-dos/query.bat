@echo off

Reg QUERY NKLM\sOFTWARE\mICROSOFT\updates\WINDOWS XP\sp2\Q322011 /v flag >nul

goto %ERRORLEVEL%

:1

goto 3

:0 

echo zero

end

:3

echo installing


:end
pause