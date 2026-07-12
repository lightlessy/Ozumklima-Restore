<%@ Language="VBScript" CodePage="65001" %>
<%
Response.Buffer = True
Response.CodePage = 65001
Response.Charset = "utf-8"
%>

<!--#INCLUDE file="dbase.asp"-->

  <%


Set rs = Server.CreateObject("ADODB.Recordset")
sql = "SELECT * FROM ayarlar WHERE kullanici ='" & request("username") & "'"
rs.open sql, baglanti,1,3
If NOT rs.EOF Then

	If request("password") = rs("sifre") Then

			Response.Cookies("ozum")("wpAfLogin") =  "true"

			Session("wpuser") = request("username")

			response.redirect "main.asp"

		else

	        	response.redirect "index.asp?lg=nousername"
		end if


ELSE

	response.redirect "index.asp?lg=nousername"



End If%>