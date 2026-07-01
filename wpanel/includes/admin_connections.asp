<%
	Set ObjConn	= Server.CreateObject("ADODB.Connection")

	StrDbName	=   "../../db/ksrtl8nkscd.mdb"
	
	StrConn = "DRIVER={Microsoft Access Driver (*.mdb)}; DBQ=" & Server.MapPath(StrDbName)

	ObjConn.Open StrConn
%>