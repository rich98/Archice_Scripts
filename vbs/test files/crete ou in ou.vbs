Set objOU1 = GetObject("LDAP://ou=OU1,dc=na,dc=fabrikam,dc=com")
Set objOU2 = objOU1.Create("Forum users", "AN&H")
objOU2.SetInfo
