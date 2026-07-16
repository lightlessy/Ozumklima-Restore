<html>
 <head>
  <title>PANEL </title> 
  <meta http-equiv="Content-Type" content="text/html; charset=utf-8">        
<link rel=stylesheet type="text/css" href="font.css">
</head>     
   
   
  <!--#include file="menu.asp"-->  
  
<div class="baslik">Popüler Ürünler</div>
<BR><BR>


<table class=font border=1 bordercolor=#E8E8E8 cellspacing=3 cellpadding=4 width=780 align=center>
 
<%


Set mods = Server.CreateObject("Adodb.Recordset")
rSQL = "Select top 10 isim,hit from products order by hit desc"
mods.Open rSQL, baglanti, 1, 3       

  Do while not mods.Eof      %>   
  


<tr class="menuoff" onmouseover="className='menuon';" onmouseout="className='menuoff';">
<td><%=mods("isim")%></td>   
<td><%=mods("hit")%></td>
   
</tr>
<%   
mods.MoveNExt
Loop
%>      
</td></tr>   
</table>  


    <center><font class=fontbig>

   <BR>         
 <BR>
<B>Ürün Hitleri </B>
           <BR><BR>
     
</font>
                            

<table class=font border=1 bordercolor=#E8E8E8 cellspacing=3 cellpadding=4 width=780 align=center>

 
<%

Set mods1 = Server.CreateObject("Adodb.Recordset")
rSQL = "Select id,isim  from  anagrup "
mods1.Open rSQL, baglanti, 1, 3     



Do while not mods1.Eof     

Set mods = Server.CreateObject("Adodb.Recordset")
rSQL = "Select anagrup,isim,hit from products  where anagrup="& mods1("id") &" group by anagrup, isim, hit order by hit desc"
mods.Open rSQL, baglanti, 1, 3       

%>  
           
<tr bgcolor=#F3F0E4><td colspan=4><B><%=mods1("isim")%></B> </td></tr> 
    
    
<% 
  k=1 
  Do while not mods.Eof      %>   
  

<tr class="menuoff" onmouseover="className='menuon';" onmouseout="className='menuoff';">
<td><%=mods("isim")%></td>   
<td><%=mods("hit")%></td>
   
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

