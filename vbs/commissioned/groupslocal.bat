if exist c:\windows\groupadmin.txt goto end

net localgroup administrators /delete forum_domain\it_dept

net localgroup administrators /add forum_domain\setup

echo it_dept removed from local adim setup added >c:\windows\groupadmin.txt

:end

