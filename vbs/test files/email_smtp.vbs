Set objEmail = CreateObject("CDO.Message")
objEmail.From = "richard.wadsworth@forumgroup.co.uk"
objEmail.To = "richard.wadsworth@forumgroup.co.uk"
objEmail.Subject = "Atl-dc-01 down" 
objEmail.Textbody = "Atl-dc-01 is no longer accessible over the network."
objEmail.Send