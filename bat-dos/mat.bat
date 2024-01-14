echo off

rasdial "vodafone gprs"

xcopy "\\forum10\matrix$\forum matrix.url" "c:\Documents and Settings\All Users\Desktop" /y

xcopy "\\forum10\matrix$\forum Matrix.ico" c:\windows /y

:end

rasdial "vodafone gprs" /disconnect

exit