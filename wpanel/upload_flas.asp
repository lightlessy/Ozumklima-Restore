<%@ Language="VBScript" CodePage="65001" %>
<%
Response.Buffer = True
Response.CodePage = 65001
Response.Charset = "utf-8"
%>

<!--#INCLUDE file="dbase.asp"-->

<%
StrUploadAdminPathName = "../slayt/"
Set Upload = Server.CreateObject("Persits.Upload")



Upload.Save Server.MapPath(StrUploadAdminPathName)

		Set File1 = Upload.Files("resim")
	       adres = Upload.Form("adres")

		'##---

		Set ObjFso = CreateObject("Scripting.FileSystemObject")

			Set ObjFileObject1 = ObjFso.GetFile(Server.MapPath(StrUploadAdminPathName) & "\" & File1.FileName)


	SQL = "INSERT INTO flas(adres, resim) Values('" & adres & "','" & ObjFileObject1.Name & "' )"
				baglanti.Execute(SQL)

			Set ObjFileObject1 = Nothing

		Set ObjFso = Nothing

		'##---

		Set File1 = Nothing

	Set Upload = Nothing

	Response.Redirect "slayt_sil.asp"
	Response.End

	Set baglanti = Nothing
%>