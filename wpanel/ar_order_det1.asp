<%@ Language="VBScript" CodePage="65001" %>
<%
Response.Buffer = True
Response.CodePage = 65001
Response.Charset = "utf-8"
%>


<html>
<head>
<title>Detaylar</title>

<link rel=stylesheet type="text/css" href="font.css">

<!--#INCLUDE file="dbase.asp"-->


<%


set mods = Server.CreateObJect("ADODB.RecordSet")
Sorgula = "Select * From arsiparis where kod = '"& request.querystring("id") &"' "
mods.open Sorgula,baglanti,1,3 %>

         <BR><BR>

<table bgcolor=#F3F3F3 cellpadding=5 cellspacing=5 align=center  border=0 width=500 class=font>

 <tr bgcolor=#FFFFFF>
 <td colspan=2 height=25>  <font color=#BF0000>
Sipariş Kodu: <B><%=mods("kod")%></B>   -  <%=mods("tarih")%>
<BR>
Kargo Takip No:  <B><%=mods("takip")%></B>
</td></tr>

<tr bgcolor=#FCFCFC><td height=25 colspan=2>
<%
Set Rsp = Server.CreateObject("Adodb.Recordset")
rSQL = "Select * from uyeler where  uid =" & mods("uid") & ""
Rsp.Open rSQL, baglanti, 1, 3  %>


<% If Rsp.EOF And Rsp.BOF Then %>

<font class=fy>
<font color=#CE0000><B>Müşteri Bulunamadı !</B>

<% else %>

<div style="padding-top:2px; padding-left:3px; padding-right:6px; padding-bottom:0px; line-height:18px ">



<table bgcolor=#FFFFFF cellpadding=0 cellspacing=0 align=left class=font border=0 bordercolor=#ACACAC width=385>
<tr bgcolor=#F3F3F3><td height=35 colspan=2><B>TESLİMAT ADRESİ</B></td></tr>
<tr><td height=25 width=70>  Ad Soyad:</td><td><%=Rsp("ad")%>&nbsp;<%=Rsp("soyad")%></td></tr>
<tr><td height=25>  Adres:  </td><td> <%=Rsp("adres")%> / <%=Rsp("sehir")%></td></tr>
<tr><td height=25>   Tel:  </td><td><%=Rsp("tel")%></td></tr>
<tr><td height=25>  TC No:  </td><td><%=Rsp("tcno")%><BR><BR></td></tr></table>




  <BR><BR>

<table bgcolor=#FFFFFF cellpadding=0 cellspacing=0 align=left class=font border=0 bordercolor=#ACACAC width=385>

<tr><td height=20 colspan=2>


<table class=font  border=0 cellspacing=4 cellpadding=6 width=780>
 <tr bgcolor=#F3F3F3>
<td height=19><B>Ürünler:</B></td>
<td width=55><B>Fiyat:</B></td>
<td width=55><B>Miktar:</B></td>
<td align=right><B>Toplam:</B></td>
</tr>


<%Do while not mods.Eof  %>


<%
set sorad = Server.CreateObJect("ADODB.RecordSet")
Sorgula = "Select * From products where AffiliateID = "& mods("product") &"  "
sorad.open Sorgula,baglanti,1,3


set sorad = Server.CreateObJect("ADODB.RecordSet")
Sorgula = "Select * From products where AffiliateID = "& mods("product") &"  "
sorad.open Sorgula,baglanti,1,3
if not sorad.eof then

set sorface = Server.CreateObJect("ADODB.RecordSet")
Sorgula = "Select * From surface where id = "& sorad("surface") &"  "
sorface.open Sorgula,baglanti,1,3

end if




if sorad.eof then %>
<BR><BR><font color=#A40000><B>Ürüne Ulaşılamıyor</B> </font><BR><BR>
<% else


set sorface = Server.CreateObJect("ADODB.RecordSet")
Sorgula = "Select * From surface where grup = "& mods("product") &"  "
sorface.open Sorgula,baglanti,1,3

%>
<tr>
<td height=20 valign=top>
<B><%=sorad("isim")%></b>
 </td>
<td valign=top align=center><%=FormatNumber(sorad("fiyat"),2)%>  TL</td>
<td valign=top align=center>    <%=mods("adet")  %>  Adet</td>
<td valign=top align=right> <%urunt=sorad("fiyat")*mods("adet")%> <%=FormatNumber(urunt,2)%> TL </td> </tr>




<% end if %>


<%
mods.MoveNExt
Loop
%>


<%
set modss = Server.CreateObJect("ADODB.RecordSet")
Sorgula = "Select * From arsiparis where kod = '"& request.querystring("id") &"' "
modss.open Sorgula,baglanti,1,3
%>




</table>

</td></tr>




</table>


<% end if %>
