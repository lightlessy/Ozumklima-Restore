<%
Response.Buffer = True
Response.CodePage = 65001
Response.Charset = "utf-8"
If Session.CodePage <> 65001 Then Session.CodePage = 65001

Set baglanti = Server.CreateObject("ADODB.Connection")
baglanti.Open "driver={microsoft access driver (*.mdb)}; dbq=" & Server.MapPath("../db/TRvd678577.mdb")
%>
