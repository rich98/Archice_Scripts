Set objOU1 = GetObject("LDAP://ou=forum users,dc=rcwadsworth,dc=co,dc=uk")
Set objOU2 = objOU1.Create("organizationalUnit", "ou=OU2")
Set objOU3 = objOU1.Create("organizationalUnit", "ou=OU3")

objOU2.SetInfo
objOU3.SetInfo