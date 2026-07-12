<%@ Language="VBScript" CodePage="65001" %>
<%
Response.Buffer = True
Response.CodePage = 65001
Response.Charset = "utf-8"
%>

<html>
<head>
<script language=javascript>

function CloseWindow()

{

window.open('','_self','');

window.close();

}

</script>
<title>Foto Güncelle</title>
<meta http-eqiv="Content-Type" content="windows-1254">
<!--#INCLUDE FILE="menu.asp"-->



<%
StrUploadAdminPathName = "../galeri/"
Set Upload = Server.CreateObject("Persits.Upload")

Upload.ProgressID = Request.QueryString("PID")
Upload.Save Server.MapPath(StrUploadAdminPathName)

		Set File1 = Upload.Files("foto")

		'##---

		Set ObjFso = CreateObject("Scripting.FileSystemObject")

			Set ObjFileObject1 = ObjFso.GetFile(Server.MapPath(StrUploadAdminPathName) & "\" & File1.FileName)

			UpdateSQL = "update galery set "
	       		UpdateSQL = UpdateSQL & "foto='"& ObjFileObject1.Name &"' WHERE id=" & Session("id")
		     	baglanti.Execute(UpdateSQL)

		Set ObjFileObject1 = Nothing

		Set ObjFso = Nothing

		'##---

%>
       <%
	Set File1 = Nothing
	Set Upload = Nothing
	Set ObjConn = Nothing   %>


<BR><BR><BR><CENTER>
<font face=arial size=4 color=#0080C0>
<b>Fotoğraf Güncellenmiştir.
<BR><BR>
<a href="galeri_edit.asp?gal_kat=<%=Session("gal_kat")%>">DEVAM ETMEK İÇİN TIKLAYINIZ >>></a>
    </b>
</CENTER>

