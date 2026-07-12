<%@ Language="VBScript" CodePage="65001" %>
<%
Response.Buffer = True
Response.CodePage = 65001
Response.Charset = "utf-8"
%>

<!--#INCLUDE file="dbase.asp"-->
<%

if request.querystring("kald") <> "" then

set ilg = Server.CreateObJect("ADODB.RecordSet")
Sorgula = "Select * From appgal where id ="& Request("rid") &" "
ilg.open Sorgula,baglanti,1,3

ilg.Delete
ilg.Update

Response.Redirect Request.ServerVariables("HTTP_REFERER")


else



toplam = Cint(Request.Form("toplam"))


for i = 1 to toplam
if Request.Form("Veri"&i) <> "" then
Set Rs = Server.CreateObject("Adodb.Recordset")
rSQL = "Select * from appgal"
Rs.Open rSQL, baglanti, 1, 3

Set Rs1 = Server.CreateObject("Adodb.Recordset")
rSQL1 = "Select * from appgal where urun="&Request.Form("Veri"&i)&" and foto="&request.form("fid")&"  "
Rs1.Open rSQL1, baglanti, 1, 3

if Rs1.eof then

Rs.AddNew
Rs("urun") = Request.Form("Veri"&i)
Rs("foto") = request.form("fid")
Rs.Update

end if

rs.close
set rs = nothing
end if
Next

baglanti.close
set baglanti = nothing

Response.Redirect Request.ServerVariables("HTTP_REFERER")


end if

%>

