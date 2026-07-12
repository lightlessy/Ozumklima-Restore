<%@ Language="VBScript" CodePage="65001" %>
<%
Response.Buffer = True
Response.CodePage = 65001
Response.Charset = "utf-8"
%>
<!--#INCLUDE file="dbase.asp"-->

<body>





<%

sil=Request.querystring("sil")
id=Request.querystring("id")



Set Rs = Server.CreateObject("Adodb.Recordset")
rSQL = "Select * from  products where AffiliateID = "& id &" "
Rs.Open rSQL,baglanti,1,3

Rs("uretici") = ""

Rs.Update
rs.close
set rs = nothing

Response.Redirect Request.ServerVariables("HTTP_REFERER")

%>




</body>
</html>
