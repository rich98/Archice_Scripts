if exist c:\windows\hid_disabled.txt goto end

net stop HidServ

sc config HidServ start= disabled

echo hid_disabled.txt >c:\windows

:end

pause