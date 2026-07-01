<%

set baglanti=server.createobject("adodb.connection")
baglanti.open "driver={microsoft access driver (*.mdb)}; dbq=" &server.mappath("../db/TRvd678577.mdb")


level2=""
SCRIPT_NAME = LCase(Request.ServerVariables("SCRIPT_NAME"))
if Instr(SCRIPT_NAME, "")>0 then 
	level2=""                                                                                                                                 
else
	level2=""
end if
                    


if Session("wpAfLogin") <> "true" then
	response.redirect level2 & "index.asp"       
		

	end if

%>


