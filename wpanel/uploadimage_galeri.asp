<%@ Language="VBScript" CodePage="65001" %>
<%
Response.Buffer = True
Response.CodePage = 65001
Response.Charset = "utf-8"
%>

<!--#INCLUDE file="menu.asp"-->

<%
StrUploadAdminPathName = "../galeri/"
Set Upload = Server.CreateObject("Persits.Upload")

Upload.ProgressID = Request.QueryString("PID")
Upload.Save Server.MapPath(StrUploadAdminPathName)




For Each File in Upload.Files

SQL = "INSERT INTO galery(gal_kat,  foto) Values(0,'" & File.ExtractFileName & "')"
baglanti.Execute(SQL)


Next

	Response.Redirect "ok.asp"
	Response.End

%>

