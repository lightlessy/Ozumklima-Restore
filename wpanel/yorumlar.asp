<%@ Language="VBScript" CodePage="65001" %>
<%
Response.Buffer = True
Response.CodePage = 65001
Response.Charset = "utf-8"

Function FixMojibake(ByVal value)
  Dim textValue
  textValue = CStr(value & "")

  textValue = Replace(textValue, "Ã‡", "Ç")
  textValue = Replace(textValue, "Ã§", "ç")
  textValue = Replace(textValue, "Ã–", "Ö")
  textValue = Replace(textValue, "Ã¶", "ö")
  textValue = Replace(textValue, "Ãœ", "Ü")
  textValue = Replace(textValue, "Ã¼", "ü")
  textValue = Replace(textValue, "Ä°", "İ")
  textValue = Replace(textValue, "Ä±", "ı")
  textValue = Replace(textValue, "Äž", "Ğ")
  textValue = Replace(textValue, "ÄŸ", "ğ")
  textValue = Replace(textValue, "Åž", "Ş")
  textValue = Replace(textValue, "ÅŸ", "ş")

  textValue = Replace(textValue, "Ã", "Ç")
  textValue = Replace(textValue, "Ã", "Ö")
  textValue = Replace(textValue, "Ã", "Ü")
  textValue = Replace(textValue, "Ã", "İ")
  textValue = Replace(textValue, "Ã", "Ş")
  textValue = Replace(textValue, "Ã¾", "ş")
  textValue = Replace(textValue, "Ã°", "ğ")

  textValue = Replace(textValue, "â€”", "—")
  textValue = Replace(textValue, "â€“", "–")
  textValue = Replace(textValue, "â€™", "’")
  textValue = Replace(textValue, "â€œ", "“")
  textValue = Replace(textValue, "â€", "”")

  FixMojibake = textValue
End Function
%>

<!--#INCLUDE file="menu.asp"-->

<body>

<div class="baslik">Yorumlar</div>

<%
intGecerliSayfa = Request.QueryString("sayfa")
If intGecerliSayfa = "" Or IsNumeric(intGecerliSayfa) = False Then intGecerliSayfa = 1
%>

<br><br><center>

<% If Request("delete") <> "" Then %>
<%
Set Rs = Server.CreateObject("Adodb.Recordset")
rSQL = "SELECT * FROM yorumlar WHERE id=" & Request.Form("id")
Rs.Open rSQL, baglanti, 1, 3
Rs.Delete
Rs.Update
Rs.Close
Set Rs = Nothing
Response.Redirect Request.ServerVariables("HTTP_REFERER")
%>
<% End If %>

<% If Request.QueryString("update") <> "" Then %>
<%
Set Rs = Server.CreateObject("Adodb.Recordset")
rSQL = "SELECT * FROM yorumlar WHERE id=" & Request("sid")
Rs.Open rSQL, baglanti, 1, 3
Rs("onay") = 1
Rs.Update
Rs.Close
Set Rs = Nothing
Response.Redirect Request.ServerVariables("HTTP_REFERER")
%>
<% End If %>

<% If Request.Form("cevapla") <> "" Then %>
<%
Set Rs = Server.CreateObject("Adodb.Recordset")
rSQL = "SELECT * FROM yorumlar WHERE id=" & Request("id")
Rs.Open rSQL, baglanti, 1, 3
Rs("onay") = 1
Rs("cevap") = Request.Form("cevap")
Rs.Update
Rs.Close
Set Rs = Nothing
Response.Redirect Request.ServerVariables("HTTP_REFERER")
%>
<% End If %>

<%
sid = 0
If IsNumeric(Request("sid")) Then sid = CInt(Request("sid"))
%>

<br>

<%
Set object = Server.CreateObject("ADODB.RecordSet")
Sorgula = "SELECT * FROM yorumlar ORDER BY id DESC"
object.Open Sorgula, baglanti, 1, 3

If Not object.EOF Then
%>

<table border="1" bordercolor="#E6E6E6" bgcolor="#FFFFFF" cellpadding="3" cellspacing="3" width="980" class="font">
<tr bgcolor="#F3F0E4" height="30">
<td align="center"><b>Yorum</b></td>
<td align="center"><font color="#729A25"><b>Cevapla</b></font></td>
<td align="center"><font color="#0080C0"><b>Onay</b></font></td>
<td align="center"><font color="#FF0000"><b>SİL</b></font></td>
</tr>

<%
i = 1
object.PageSize = 20
object.AbsolutePage = intGecerliSayfa
intSayfaSayisi = object.PageCount

For b = 1 To object.PageSize
  If object.EOF Then Exit For
%>

<form method="post" action="yorumlar.asp">
<tr <% If object("onay") = 1 Then %>bgcolor="#F3F3F3"<% End If %>>

<td width="444"><%=FixMojibake(object("yorum"))%>
<br><br>

<% If Request("cevap") = 1 And sid = object("id") Then %>
<form action="yorumlar.asp" method="post">
<textarea style="width:400px;height:90px" name="cevap"><%=FixMojibake(object("cevap"))%></textarea>
<input type="hidden" name="id" value="<%=sid%>">
<input type="hidden" name="cevapla" value="1">
<input type="submit" value="Gönder">
</form>
<% End If %>
</td>

<td><a href="yorumlar.asp?cevap=1&sid=<%=object("id")%>"><font color="#729A25"><b>CEVAPLA</b></font></a></td>

<td align="center" width="100">
<% If object("onay") = 1 Then %>
  <b>Yayında</b>
<% Else %>
  <a href="yorumlar.asp?update=1&sid=<%=object("id")%>"><font color="#0080C0"><b>ONAYLA</b></font></a>
<% End If %>
</td>

<td align="center">
<form action="yorumlar.asp" method="post">
<input type="hidden" name="id" value="<%=object("id")%>">
<input type="submit" value="SİL" name="delete" onclick="return confirm('Kayıt silinecek. Onaylıyor musunuz?')">
</form>
</td>
</tr>

<%
object.MoveNext
Next
%>

</table>
<br><br>

<table width="900" class="font" border="1" bordercolor="#A3D1D1"><tr><td align="center">

<%
Aralik = 2
intSayfaAltSiniri = intGecerliSayfa - Aralik
intSayfaUstSiniri = intGecerliSayfa + Aralik
intSolGrupSayisi = 1 + Aralik
intSagAltSinir = intSayfaSayisi - Aralik

If intSayfaUstSiniri > intSayfaSayisi Then
  intSayfaAltSiniri = intSayfaSayisi - (2 * Aralik)
  intSayfaUstSiniri = intSayfaSayisi
End If

If intSayfaAltSiniri <= 0 Then
  intSayfaAltSiniri = 1
  intSayfaUstSiniri = intSayfaAltSiniri + (2 * Aralik)
  If intSayfaUstSiniri >= intSayfaSayisi Then intSayfaUstSiniri = intSayfaSayisi
End If

If intSolGrupSayisi >= intSayfaAltSiniri Then intSolGrupSayisi = intSayfaAltSiniri - 1
If intSagAltSinir <= intSayfaUstSiniri Then intSagAltSinir = intSayfaUstSiniri + 1
%>

<% If intGecerliSayfa > 1 Then %>
<a href="?sayfa=1" title="İlk Sayfa"><font color="#4E4E4E">[««]</font></a>
<a href="?sayfa=<%=intGecerliSayfa-1%>" title="Önceki Sayfa"><font color="#4E4E4E">[«]</font></a>
<% End If %>

<%
For i = 1 To intSolGrupSayisi
  If CInt(i) = CInt(intGecerliSayfa) Then
    Response.Write " <b>" & i & ".</b>"
  Else
    Response.Write " <a href=""?sayfa=" & i & """ title=""" & i & ". Sayfa""><font color=""#4E4E4E"">" & i & ".</font></a>"
  End If
Next

If intSayfaAltSiniri - intSolGrupSayisi > 1 Then Response.Write "<b>. . .</b>"

For i = intSayfaAltSiniri To intSayfaUstSiniri
  If CInt(i) = CInt(intGecerliSayfa) Then
    Response.Write " <b>" & i & ".</b>"
  Else
    Response.Write " <a href=""?sayfa=" & i & """ title=""" & i & ". Sayfa""><font color=""#4E4E4E"">" & i & ".</font></a>"
  End If
Next

If intSagAltSinir - intSayfaUstSiniri > 1 Then Response.Write "<b>. . .</b>"

For i = intSagAltSinir To intSayfaSayisi
  If CInt(i) = CInt(intGecerliSayfa) Then
    Response.Write " <b>" & i & ".</b>"
  Else
    Response.Write " <a href=""?sayfa=" & i & """ title=""" & i & ". Sayfa""><font color=""#4E4E4E"">" & i & ".</font></a>"
  End If
Next
%>

<% If CInt(intGecerliSayfa) <> CInt(intSayfaSayisi) Then %>
<a href="?sayfa=<%=intGecerliSayfa+1%>" title="Sonraki Sayfa"><font color="#4E4E4E">[»]</font></a>
<a href="?sayfa=<%=intSayfaSayisi%>" title="Son Sayfa"><font color="#4E4E4E">[»»]</font></a>
<% End If %>

</td></tr></table>

<%
object.Close
Set object = Nothing
End If
%>

<br><br>
