strComputer = "\\lt-r-c-wads"
Set objGroup = GetObject("windows://" & strComputer & "/Administrators,group")
Set objUser = GetObject("Windows://" & strComputer & "/(forum_domain/oliver.white,user")
objGroup.Add(objUser.ADsPath)
