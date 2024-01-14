REM ***** stonegate program copy *****

if exist c:\windows\stonesoftbak\sgclient-2.2.1-624.exe goto end

mkdir c:\windows\stonesoftbak

xcopy \\forum10\stonesoftbak$\*.* c:\windows\stonesoftbak /y

:end



