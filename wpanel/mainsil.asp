<html>
 <head>
  <title>PANEL </title> 
  <meta http-equiv="Content-Type" content="text/html; charset=utf-8">        
<link rel=stylesheet type="text/css" href="font.css">
</head>     

   
  <!--#include file="menu.asp"-->  
 
 <div class="baslik"> Ana Kategori Düzenle</div> 

 <BR><BR>
  
<%


if Request.Querystring("psil2")=1 then 


Set Rs = Server.CreateObject("Adodb.Recordset")
rSQL = "Select * from anagrup  where id = "& Request("id")
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
rSQL = "Select * from anagrup where id = "& Request.Form("id")
Rs.Open rSQL, baglanti, 1, 3 

Rs("isim") = Request("isim")
Rs("sira") = Request("sira") 
Rs("durum") = Request("durum")
Rs("keyw") = Request("keyw")
Rs("descr") = Request("descr")

Rs.Update 
rs.close
set rs = nothing 
Response.Redirect Request.ServerVariables("HTTP_REFERER")            

 end if %>

                    

<BR><BR>   



<%    if   request.querystring("edit") = "yes"  then %>     

 
<%

Set mods = Server.CreateObject("Adodb.Recordset")
rSQL = "Select * from anagrup where id = "& Request("id")
mods.Open rSQL, baglanti, 1, 3 

%>  
       


 <form method="POST" action="mainsil.asp?edit=yes&id=<%=mods("id")%>"   name="myForm" id="myForm">   
    

<table class=font border=1 bordercolor=#E8E8E8 bgcolor=#FFFFFF cellspacing=3 cellpadding=4 width=380 align=center>


<tr> 
<td>Sıra:        <input type="text"  name="sira"   size="2" value="<%=mods("sira")%>"> </td>  
<td> İsim:       <input type="text"  name="isim"   size="30" value="<%=mods("isim")%>"> </td>    
</tr>
<tr><td colspan=2>  
<input type="checkbox" name="durum"  value="1" <% if mods("durum")="1" then %> checked <% end if %>> Kategori Aktif

</td></tr>
<tr>
<td colspan=2>Keywords: <BR><TEXTAREA   style="WIDTH: 550px; HEIGHT: 59px"  name="keyw"><%=mods("keyw")%></TEXTAREA> </td>    
</tr>   
<tr>
<td colspan=2>Description: <BR><TEXTAREA   style="WIDTH: 550px; HEIGHT: 59px"  name="descr"><%=mods("descr")%></TEXTAREA> </td>    
</tr>   

 <input type="hidden" name="id" value="<%=mods("id")%>">  
 
<tr><td colspan=2><input type="submit" name="update" value="Düzenle"></td>   
</form>                                                                                                            
</tr></table>
                      

<% else %>




<table class=font border=1 bordercolor=#E8E8E8 bgcolor=#FFFFFF cellspacing=3 cellpadding=4 width=380 align=center>


<%

Set mods = Server.CreateObject("Adodb.Recordset")
rSQL = "Select * from anagrup  order by sira desc"
mods.Open rSQL, baglanti, 1, 3 
 
Do while not mods.Eof      
%>  
<tr class="menuoff" onmouseover="className='menuon';" onmouseout="className='menuoff';">
<td>           
<%=mods("isim")%>   </td><td align=center><a href="?edit=yes&id=<%=mods("id")%>"><img src="images/edit.png" width=20></a><td align=center> <a href="?psil2=1&id=<%=mods("id")%>" onclick="return confirm('Kayıt silinecek. Onaylıyor Musunuz?')"><img src="images/del.png" width=20></a>
</td>      
</tr>
<%  
mods.MoveNExt
Loop

%>  
</table>   


<% end if %>

