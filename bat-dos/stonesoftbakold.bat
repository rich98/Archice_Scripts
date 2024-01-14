REM ***** stonegate program copy *****

if exist c:\stonesoftbak\client-certificate.pem goto end

mkdir c:\windows\stonesoftbak

xcopy \\forum10\stonesoftbak$\*.* c:\windows\stonesoftbak



:end



