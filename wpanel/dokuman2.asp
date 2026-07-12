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


<!--#INCLUDE file="dbase.asp"-->





<%
StrUploadAdminPathName = "../urunler/"
Set Upload = Server.CreateObject("Persits.Upload")



Upload.ProgressID = Request.QueryString("PID")

Upload.OverwriteFiles = False

Upload.Save Server.MapPath(StrUploadAdminPathName)

		Set File1 = Upload.Files("dokuman")
	       	id = Upload.Form("id")
	g = Upload.Form("g")
	tip = Upload.Form("tip")

		Set ObjFso = CreateObject("Scripting.FileSystemObject")
	Set ObjFileObject1 = ObjFso.GetFile(Server.MapPath(StrUploadAdminPathName) & "\" & File1.FileName)

	UpdateSQL = "update products set "
				UpdateSQL = UpdateSQL & "dokuman='"& ObjFileObject1.Name &"' WHERE AffiliateID=" & id
				baglanti.Execute(UpdateSQL)

%>


<meta http-equiv="refresh"
content="1; URL=edits.asp?id=<%=id%>&g=<%=g%>&tip=<%=tip%>">



                 <BR><BR><BR><CENTER>
<font face=arial size=4 color=#0080C0>
<b>Döküman Yüklendi... Lütfen Bekleyiniz.</b>
</CENTER>

