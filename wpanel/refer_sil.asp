   
<!--#INCLUDE file="menu.asp"-->

  <body>  

<div class="baslik">Referans Düzenle </div>
<BR><BR>



<%if request("delete")<>"" then  %>

<%


Set Rs = Server.CreateObject("Adodb.Recordset")
rSQL = "Select * from refers where id = "& Request.Form("id")
Rs.Open rSQL, baglanti, 1, 3 
Rs.Delete ' Sil
Rs.Update ' Güncelle

rs.close
set rs = nothing 

Response.Redirect Request.ServerVariables("HTTP_REFERER")
%>        
<% end if %>   


<%if request("update")<>"" then  %>

<%

Set Rs = Server.CreateObject("Adodb.Recordset")
rSQL = "Select * from refers where id = "& Request.Form("id")
Rs.Open rSQL, baglanti, 1, 3 

Rs("isim") = Request("isim")          
Rs("link") = Request("link") 
Rs("sira") = Request("sira")    

Rs.Update 

rs.close
set rs = nothing 

Response.Redirect Request.ServerVariables("HTTP_REFERER")
%>        
<% end if %>

<% 
edit = Request("edit")
sid = Cint(Request("sid")) %>                         




      <%    if edit = "yes" then 

  sid = request.querystring("sid")   

set object = Server.CreateObJect("ADODB.RecordSet")
Sorgula = "Select * From refers where id = "& sid
object.open Sorgula,baglanti,1,3             

%>      


 <form method="post" name="f" action="refer_sil.asp"> 
   	 
<table border="1" bordercolor="#E6E6E6" bgcolor=#FFFFFF cellpadding="3" cellspacing="3" width=700 class=font align=center>   

<tr> <td><b>Ýsim:</b> </td><td><input type="text" name="isim"   value="<%=object("isim")%>" size="95"></td></tr> 
          
<tr> <td><b>Link:</b> </td><td><input type="text" name="link"   value="<%=object("link")%>" size="95"></td></tr> 

<tr> <td><b>Sýralama:</b> </td><td><input type="text" name="sira"   value="<%=object("sira")%>" size="5"></td></tr>
   

<input type="hidden" name="id" value="<%=sid%>"> 
<tr><td align=center colspan=3>
<input type="submit" name="update" value="Güncelle">

</td>      </form>
</div>        
  
</tr>
<tr><td colspan=2> 
<a href="refer_sil.asp"><font color="#3E3E3E" size=2>  <B>« Referans Listesi  </B> </font></a>   </td></tr>
   </tr></table>  
   
 





<% else %>

<%  

set object = Server.CreateObJect("ADODB.RecordSet")
Sorgula = "Select * From refers order by sira desc"
object.open Sorgula,baglanti,1,3  

%>      


<table border="1" bordercolor="#E6E6E6" bgcolor=#FFFFFF cellpadding="3" cellspacing="3" width=700 class=font align=center>
<tr bgcolor=#F3F0E4 height=30>    
                                                    
<td align=center><b>Ýsim</b></td>
<td align=center><b>Link</b></td>
<td align=center><b>Görsel</b></td>
<td align=center><font color=#0080C0><b>Düzenle</b></td><td align=center><font color=#FF0000><b>SÝL</b></td></tr>      
  
  	 
 
<%
Do while not object.Eof                                                                                                                                                                                                                  
%>   
  

<tr class="menuoff" onmouseover="className='menuon';" onmouseout="className='menuoff';">
 <td><b><%=object("isim")%></td>  
  <td><b><%=object("link")%></td>  

 <td align=center><img src="../ref/<%=object("resim")%>" width=150></td>  

<td align=center>
<a href="refer_sil.asp?edit=yes&sid=<%=object("id")%>"><img src="images/edit.png" width=20></a>
</td>

<form action="refer_sil.asp" method="post" >
<td align=center>    
<input type="hidden" name="id" value="<%=object("id")%>">      
<input type="hidden" name="delete" value="1">        
<input type="image" src="images/del.png"  onclick="return confirm('Kaydý Silmek Ýstiyor Musunuz?')"></td></form>


</div>        




<%  
object.MoveNExt
Loop
%>
             
      

   </tr></table>  
                                   

<%
object.close
set object = nothing

%>             
             
      
      <% end if %>     
      <BR><BR>


 
 