<%@ Language="VBScript" CodePage="65001" %>
<%
Response.Buffer = True
Response.CodePage = 65001
Response.Charset = "utf-8"
%>

<html>
<head>

<title>Foto Güncelle</title>
<meta http-eqiv="Content-Type" content="windows-1254">

    <!--#INCLUDE file="dbase.asp"-->

      <!--#include file="menu.asp"-->



<%

StrUploadAdminPathName = "../urunler/"
Set Upload = Server.CreateObject("Persits.Upload")

Upload.OverwriteFiles = False

Upload.ProgressID = Request.QueryString("PID")
Upload.Save Server.MapPath(StrUploadAdminPathName)

	Set File1 = Upload.Files("foto")

	Set ObjFso = CreateObject("Scripting.FileSystemObject")
	Set ObjFileObject1 = ObjFso.GetFile(Server.MapPath(StrUploadAdminPathName) & "\" & File1.FileName)


			UpdateSQL = "update products set "
	       		UpdateSQL = UpdateSQL & "foto"& Upload.form("n") &"='"& ObjFileObject1.Name &"' WHERE AffiliateID=" & Upload.form("id") &""
		     	baglanti.Execute(UpdateSQL)



%>

   <meta http-equiv="refresh"
content="1; URL=edits.asp?id=<%=Upload.form("id")%>&g=<%=Upload.form("g")%>&tip=<%=Upload.form("tip")%>">



                 <BR><BR><BR><CENTER>
<font face=arial size=4 color=#0080C0>
<b>Resim Yüklendi... Lütfen Bekleyiniz.</b>



</CENTER>

