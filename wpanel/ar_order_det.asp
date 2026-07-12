<%@ Language="VBScript" CodePage="65001" %>
<%
Response.Buffer = True
Response.CodePage = 65001
Response.Charset = "utf-8"
%>
<html>
<head>
<title>Tamamlanmış Siparişler</title>

  <!--#include file="menu.asp"-->


                 <link rel="stylesheet" href="../css/lity.css">
                  <script src="../js/lity.js"></script>



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


  <div class="baslik">Tamamlanmış Siparişler </div>   <BR><BR>


<TABLE border=0 class=onay cellspacing=3 cellpadding=4 align=center><TR>
<TD valign=top>

<%
set mods1 = Server.CreateObJect("ADODB.RecordSet")
Sorgula = "Select kod, gun, tarih,onay From  arsiparis where gun LIKE '%"& request.querystring("gun") &"%'  and onay=1 group by kod, gun, tarih,onay order by tarih desc"
mods1.open Sorgula,baglanti,1,3

%>


<TD valign=top>
<font color=#0080C0 size=3>
<B><%=request.querystring("gun") %></B>   Tarihli <font color=#B90000><B><%=mods1.recordcount %></B> </font> Sipariş </font>

<BR>

<% Do while not mods1.Eof  %>

<BR>
<table bgcolor=#f3f7f8 cellpadding=2 cellspacing=2 align=center class=font border=0 width=500>
<tr bgcolor=#FFFFFF>
<td colspan=2 class=onay>

<div id="inline" style="background:#fff; padding:5px; font-family:Open Sans; font-size:14px; color:#373737; width:1000px; height:550px  " class="lity-hide">
<IFRAME frameBorder=0 scrolling=no width=1000 height=600 src="ar_order_det1.asp?id=<%=mods1("kod")%>" ></IFRAME>
</div>


<a href="#inline" data-lity><font color=#4E4E4E> <B> <%=mods1("kod")%> - Sipariş Detayları  »</B> </font>  </a>
</td></tr>


</table>

<%
mods1.MoveNExt
Loop
%>


</TD></TR>

<TR><TD align=center colspan=2>
<a href="ar_orders.asp"><font color=#0080C0 size=2> Listeye Dön  «  </font></a>
</TD></TR></TABLE>


<BR>




</body>
</html>