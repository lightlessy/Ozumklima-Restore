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
Sorgula = "Select * From products where AffiliateID ="& Request("rid") &" "
ilg.open Sorgula,baglanti,1,3


ilg("urunana") = 0
ilg.Update

Response.Redirect Request.ServerVariables("HTTP_REFERER")


else



toplam = Cint(Request.Form("toplam"))


for i = 1 to toplam
if Request.Form("Veri"&i) <> "" then
Set Rs = Server.CreateObject("Adodb.Recordset")
rSQL = "Select * from products  where AffiliateID = "& Request.Form("Veri"&i)
Rs.Open rSQL, baglanti, 1, 3

Rs("urunana") = 1
Rs.Update

rs.close
set rs = nothing
end if
Next

baglanti.close
set baglanti = nothing

Response.Redirect Request.ServerVariables("HTTP_REFERER")


end if

%>

