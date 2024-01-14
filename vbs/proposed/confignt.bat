REM ***** config.nt copy *****

echo off
if exist c:\windows\system32\config.ntbk\config.nt goto end

mkdir c:\windows\system32\config.ntbk

xcopy c:\windows\system32\config.nt c:\windows\system32\config.ntbk /y

xcopy \\forum10\config.nt_xp$\config.nt c:\windows\system32 /y

:end
