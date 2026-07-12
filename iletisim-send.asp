<%@ Language="VBScript" CodePage="65001" %>
<%
Response.Buffer = True
Session.CodePage = 65001
Session.LCID = 1055
Response.CodePage = 65001
Response.Charset = "utf-8"
Response.ContentType = "text/html"
%>

<!--#include file="config/mail-config.asp"-->
<!DOCTYPE html>
<html>
<head>
<meta charset="utf-8">
<title>Mesaj Gönderimi</title>
</head>

<body style="font-family:Arial; font-size:15px; padding:40px;">

<%
On Error Resume Next

isim = Trim(Request.Form("isim"))
mail = Trim(Request.Form("mail"))
tel = Trim(Request.Form("tel"))
formMesaj = Trim(Request.Form("mes"))

If isim = "" Or mail = "" Or formMesaj = "" Then
%>

  <b>Eksik bilgi var.</b><br><br>
  Lütfen formu tekrar doldurun.

<%
  Response.End
End If

htmlBody = ""
htmlBody = htmlBody & "<table style='font-family:arial; font-size:14px; width:500px; border:1px solid #E6E6E6' cellspacing='6' cellpadding='5'>"
htmlBody = htmlBody & "<tr><td colspan='2'><font color='#0080C0'><b>Bize Ulaşın Formu</b></font></td></tr>"
htmlBody = htmlBody & "<tr><td>Ad Soyad:</td><td><b>" & Server.HTMLEncode(isim) & "</b></td></tr>"
htmlBody = htmlBody & "<tr><td>E-Posta:</td><td><b>" & Server.HTMLEncode(mail) & "</b></td></tr>"
htmlBody = htmlBody & "<tr><td>Telefon:</td><td><b>" & Server.HTMLEncode(tel) & "</b></td></tr>"
htmlBody = htmlBody & "<tr bgcolor='#F3F3F3'><td>Mesaj:</td><td><b>" & Replace(Server.HTMLEncode(formMesaj), vbCrLf, "<br>") & "</b></td></tr>"
htmlBody = htmlBody & "</table>"

kime = CONTACT_TO
kimden = SMTP_USER
konu = "Bize Ulaşın"
gonderenAd = CONTACT_FROM_NAME

Set mektup = Server.CreateObject("Persits.MailSender")

If Err.Number <> 0 Then
%>

  <b>Mail sistemi hatası:</b><br><br>
  Persits.MailSender sunucuda kurulu değil veya izin yok.<br><br>
  Hata: <%=Err.Description%>

<%
  Response.End
End If

mektup.Host = SMTP_HOST
mektup.Port = SMTP_PORT
mektup.IsHTML = True

mektup.From = kimden
mektup.FromName = gonderenAd
mektup.AddAddress kime
mektup.AddReplyTo mail

mektup.Subject = konu
mektup.Body = htmlBody

mektup.Username = SMTP_USER
mektup.Password = SMTP_PASS

mektup.CharSet = "utf-8"
mektup.ContentTransferEncoding = "Quoted-Printable"

Err.Clear
mektup.Send

If Err.Number <> 0 Then
%>

  <b>Mail gönderilemedi.</b><br><br>
  Hata: <%=Err.Description%><br><br>
  Muhtemel sebepler:<br>
  - SMTP şifresi yanlış<br>
  - Hosting port 587 engelliyor<br>
  - mail.ozumklima.com SMTP kabul etmiyor<br>
  - Persits kurulu ama mail auth/TLS ayarı eksik

<%
  Response.End
End If

Set mektup = Nothing
%>

<b>Mesajınız iletilmiştir. Teşekkürler.</b>

<br><br>

<script>
  setTimeout(function () {
    window.location.href = "/tesekkurler.asp";
  }, 1000);
</script>

</body>
</html>
