<%@ Language="VBScript" CodePage="65001" %>
<%
Response.Buffer = True
Response.CodePage = 65001
Response.Charset = "utf-8"
%>

<%
set baglanti=server.createobject("adodb.connection")
baglanti.open "driver={microsoft access driver (*.mdb)}; dbq=" &server.mappath("../db/TRvd678577.mdb")

acik=Request.Form("giden")


a=split(acik,",")

uzunluk=uBound(a)

%>



<% for i=0 to uzunluk %>

<% b=split(a(i),"-")

Set Rs = Server.CreateObject("Adodb.Recordset")
rSQL = "Select * from galery where  id = "& b(0)
Rs.Open rSQL, baglanti, 1, 3
Rs("sira")= b(1)
Rs.Update

next %>


