<html>
 <head>
  <title>PANEL </title> 
  <meta http-equiv="Content-Type" content="text/html; charset=utf-8">        
<link rel=stylesheet type="text/css" href="font.css">
</head>     
   
   
  <!--#include file="menu.asp"-->  
  
<div class="baslik">İndirimli Ürünler </div>
           <BR><BR>
  
<%


if Request.Querystring("psil2")=1 then 


Set Rs = Server.CreateObject("Adodb.Recordset")
rSQL = "Select * from products  where AffiliateID = "& Request("id")
Rs.Open rSQL, baglanti, 1, 3 

Rs.Delete 
Rs.Update 

rs.close
set rs = nothing 
Response.Redirect Request.ServerVariables("HTTP_REFERER")         

end if      %>   

                            

<table class=font border=1 bordercolor=#E8E8E8 cellspacing=3 cellpadding=4 width=600 align=center>

 
<%

Set mods1 = Server.CreateObject("Adodb.Recordset")
rSQL = "Select id,isim  from  anagrup "
mods1.Open rSQL, baglanti, 1, 3     



Do while not mods1.Eof     

Set mods = Server.CreateObject("Adodb.Recordset")
rSQL = "Select AffiliateID,anagrup,isim,urunana,tip,indirim from products  where anagrup="& mods1("id") &" and indirim<>0 group by anagrup, isim, AffiliateID,urunana,tip,indirim"
mods.Open rSQL, baglanti, 1, 3       

%>  
           
<tr bgcolor=#F3F0E4><td colspan=4><B><%=mods1("isim")%></B> </td></tr> 
    
    
<% 
  k=1 
  Do while not mods.Eof      %>   
  

<tr class="menuoff" onmouseover="className='menuon';" onmouseout="className='menuoff';">
<td><%=mods("isim")%></td>

<td align=center><a href="edits.asp?id=<%=mods("AffiliateID")%>&g=<%=mods("anagrup")%>&tip=<%=mods("tip")%>"><img src="images/edit.png" width=20></a>    
</tr>
<%   
  k = k + 1 
mods.MoveNExt
Loop
%>      
</td></tr>   
    

<%     

mods1.MoveNExt
Loop
%>  
</table>  

                         
<BR><BR>
&nbsp;

