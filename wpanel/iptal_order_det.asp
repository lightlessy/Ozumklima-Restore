<%@ Language="VBScript" CodePage="65001" %>
<%
Response.Buffer = True
Response.CodePage = 65001
Response.Charset = "utf-8"
%>
<html>
<head>
<title>İptal Siparişler</title>

  <!--#include file="menu.asp"-->

  <%if request.querystring("del") ="1" then  %>

<%

Set Rs = Server.CreateObject("Adodb.Recordset")
rSQL = "Select * from arsiparis where kod = '"& Request.querystring("kod") & "' "
Rs.Open rSQL, baglanti, 1, 3
Rs.Delete
Rs.Update

rs.close
set rs = nothing
Response.Redirect "ar_orders.asp"
%>
<% end if %>

</HEAD>
<body>


  <div class="baslik">İptal Siparişler </div>

  <BR><BR>


<TABLE border=0 class=onay cellspacing=3 cellpadding=4 align=center><TR>
<TD valign=top>

<%
set mods1 = Server.CreateObJect("ADODB.RecordSet")
Sorgula = "Select kod, gun, tarih,onay From  arsiparis where gun LIKE '%"& request.querystring("gun") &"%'  and onay=0 group by kod, gun, tarih,onay order by tarih desc"
mods1.open Sorgula,baglanti,1,3

%>


<TD valign=top>
<font color=#0080C0 size=3>
<B><%=request.querystring("gun") %>  </b> Tarihli <font color=#B90000><B><%=mods1.recordcount %></B> </font> Sipariş </font>

<BR>

<% Do while not mods1.Eof  %>

<BR>
<table bgcolor=#f3f7f8 cellpadding=2 cellspacing=2 align=center class=font border=0 width=500>
<tr bgcolor=#FFFFFF>
<td colspan=2 class=onay>
<a href="javascript:void(window.open('ar_order_det1.asp?id=<%=mods1("kod")%>','','width=850,height=550,left=400,top=150,resizable=no,menubar=no,centerscreen=yes,status=no,scrollbars=1'))"><font color=#4E4E4E>  <B><%=mods1("kod")%></B> »    </font>  </a>
</td></tr>


</table>

<%
mods1.MoveNExt
Loop
%>

<TR><TD align=center colspan=2 ><a href="ar_orders.asp"><font color=#0080C0 size=2>  Listeye Dön  «  </font></a>    </TD></TR></TABLE>



<BR><BR>



</body>
</html>