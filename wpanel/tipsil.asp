<%@ Language="VBScript" CodePage="65001" %>
<%
Response.Buffer = True
Response.CodePage = 65001
Response.Charset = "utf-8"
%>
<html>
 <head>
  <title>PANEL </title>
  <meta charset="utf-8">
<link rel=stylesheet type="text/css" href="font.css">
</head>


  <!--#include file="menu.asp"-->

   <div class="baslik"> Alt Kategori Düzenle</div>


<%


if Request.Querystring("psil2")=1 then

 id= request.querystring("id")

UpdateSQL = "update products set yayin=0 WHERE tip ="& id &""
baglanti.Execute(UpdateSQL)


Set Rs = Server.CreateObject("Adodb.Recordset")
rSQL = "Select * from tip  where id = "& id
Rs.Open rSQL, baglanti, 1, 3

Rs.Delete
Rs.Update

rs.close
set rs = nothing
Response.Redirect Request.ServerVariables("HTTP_REFERER")

end if      %>




<%if request("update")<>"" then  %>

<%

Set Rs = Server.CreateObject("Adodb.Recordset")
rSQL = "Select * from tip where id = "& Request.Form("id")
Rs.Open rSQL, baglanti, 1, 3

Rs("isim") = Request("isim")
Rs("sira") = Request("sira")

Rs.Update
rs.close
set rs = nothing
Response.Redirect Request.ServerVariables("HTTP_REFERER")

 end if %>



<BR><BR>



<%    if   request.querystring("edit") = "yes"  then %>


<%

Set mods = Server.CreateObject("Adodb.Recordset")
rSQL = "Select * from tip where id = "& Request("id")
mods.Open rSQL, baglanti, 1, 3

Set mods1 = Server.CreateObject("Adodb.Recordset")
rSQL = "Select id,isim  from  anagrup where id="&mods("grup")&" "
mods1.Open rSQL, baglanti, 1, 3

%>



 <form method="POST" action="tipsil.asp?edit=yes&id=<%=mods("id")%>"   name="myForm" id="myForm">


<table class=font border=1 bgcolor=#FFFFFF bordercolor=#E8E8E8 cellspacing=3 cellpadding=4 width=380 align=center>

<tr bgcolor=#F3F0E4 height=15><td colspan=2>Bağlı Olduğu Ana Kategori: <B><%=mods1("isim")%></B></td></tr>

<tr>
<td>   Sıra:     <input type="text"  name="sira"   size="3" value="<%=mods("sira")%>"> </td>
<td>       <input type="text"  name="isim"   size="30" value="<%=mods("isim")%>"> </td>
</tr>

 <input type="hidden" name="id" value="<%=mods("id")%>">

<tr><td colspan=2><input type="submit" name="update" value="Düzenle"></td>
</form>
</tr></table>


<% else %>




<table class=font border=1 bgcolor=#FFFFFF bordercolor=#E8E8E8 cellspacing=3 cellpadding=4 width=380 align=center>


<%

Set mods1 = Server.CreateObject("Adodb.Recordset")
rSQL = "Select id,isim  from  anagrup "
mods1.Open rSQL, baglanti, 1, 3

Do while not mods1.Eof

Set mods = Server.CreateObject("Adodb.Recordset")
rSQL = "Select id,sira,grup,isim from tip  where grup="& mods1("id") &" group by grup, sira, isim, id order by sira desc"
mods.Open rSQL, baglanti, 1, 3

%>

<tr bgcolor=#F3F0E4><td colspan=4><B><%=mods1("isim")%></B> </td></tr>

<% Do while not mods.Eof      %>
<tr class="menuoff" onmouseover="className='menuon';" onmouseout="className='menuoff';">
<td><%=mods("isim")%></td><td align=center><a href="?edit=yes&id=<%=mods("id")%>"><img src="images/edit.png" width=20></a><td align=center> <a href="?psil2=1&id=<%=mods("id")%>" onclick="return confirm('Kayıt silinecek. Onaylıyor Musunuz?')"><img src="images/del.png" width=20></a>
</td>
</tr>
<%
mods.MoveNExt
Loop
%>

</td></tr>

<%
mods1.MoveNExt
Loop
%>
</table>


<% end if %>

<BR><BR>&nbsp;
