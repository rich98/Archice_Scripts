cls
Echo off
echo *******press a key to stop exchange and copy exchange databases*******
pause

net stop MSExchangeSA /yes

net stop MSExchangeSA /yes

net send lt-richard-wads check services on exchange server

xcopy d:\exchsrvr\MDBDATA\*.edb d:\offline

echo ******press a key to start exchange******
pause

net start "Microsoft Exchange System Attendant"
net start "Microsoft Exchange Directory"
net start "Microsoft Exchange Information Store"
net start "Microsoft Exchange Message Transfer Agent"
net start "Microsoft Exchange Internet Mail Service"

net send lt-richard-wads check services on exchange server

:end