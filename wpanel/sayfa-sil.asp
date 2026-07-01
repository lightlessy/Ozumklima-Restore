   
<!--#INCLUDE file="menu.asp"-->  

 <div class="baslik">Ýçerik Düzenle</div>  

  <body OnLoad="initialiseWebWizRTE();">  
  
   <%	intGecerliSayfa = Request.Querystring("sayfa")
If intGecerliSayfa = "" OR IsNumeric(intGecerliSayfa) = False Then intGecerliSayfa = 1
%>
   

<BR><BR><center>
                                  




<%if request("delete")<>"" then  %>

<%


Set Rs = Server.CreateObject("Adodb.Recordset")
rSQL = "Select * from sayfalar where id = "& Request.Form("id")
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
rSQL = "Select * from sayfalar where id = "& Request.Form("id")
Rs.Open rSQL, baglanti, 1, 3 
   
Rs("isim") = Request("isim") 
Rs("ana") = Request("ana")
Rs("icerik") = Request("icerik")    
Rs("sira") = Request("sira")
Rs("yayin") = Request("yayin")    
Rs.Update 

rs.close
set rs = nothing 

Response.Redirect Request.ServerVariables("HTTP_REFERER")
%>        
<% end if %>

<% 
edit = Request("edit")
sid = Cint(Request("sid")) %>                         


<BR>


      <%    if edit = "yes" then 

  sid = request.querystring("sid")   

set object = Server.CreateObJect("ADODB.RecordSet")
Sorgula = "Select * From sayfalar where id = "& sid
object.open Sorgula,baglanti,1,3             

%>      


 <form method="post" name="myForm" action="sayfa-sil.asp"> 
   	 
<table border="1" bordercolor="#E6E6E6" bgcolor=#FFFFFF cellpadding="3" cellspacing="3" width=800 class=font>          

 <tr><td><b>Sayfa  Kategorisi </td><td  colspan=3>

<select name="ana"> 
<option value="0" <% if object("ana")=0 then %> selected <% end if %>  >Lütfen Seçiniz</option>   
<option  value="1" <% if object("ana")=1 then %> selected <% end if %> >Kurumsal</option>    
<option  value="2"  <% if object("ana")=2 then %> selected <% end if %> >Yardým</option>   
</select>  

</td></tr>	     
 
     

 
<tr> <td><b>Sayfa Adý:</b> </td><td><input type="text" name="isim"   value="<%=object("isim")%>" size="55"></td></tr>   



<tr><td colspan=2>      
	<%
strFormName = "myForm"
strTextAreaName = "icerik"
%> 
     
    <!--#include file="RTE_editor_inc.asp" --><BR>
    
 <TEXTAREA cols=100 name="icerik" id="icerik" rows=33><%=object("icerik")%></TEXTAREA>
 </td></tr>
<tr><td align=left colspan=2> <B>Sýralama:</B> &nbsp;&nbsp; <input type="text" size=3 name="sira" value="<%=object("sira")%>" onkeypress="return numbersonly(this, event)">       </td></tr>
      
<tr><td colspan=2><B>Yayýn:</B> <input type="checkbox" name="yayin"  value="1" <% if object("yayin")="1" then %> checked <% end if %>></td></tr>      

               

    <input type="hidden" name="id" value="<%=sid%>"> 
<tr><td align=center colspan=3>
<input type="submit" name="update" value="Güncelle">

</td>      </form>
</div>        
  

   </tr></table>  
   
 

<%
object.close
set object = nothing
%>   
<BR><BR>
<a href="sayfa-sil.asp"><font color=#C40000><b> « Sayfa Listesi </b>     </font></a> 
<BR><BR>          
      


                
      
      
      
<% else %>

<%  

set object = Server.CreateObJect("ADODB.RecordSet")
Sorgula = "Select * From sayfalar where ana= "& request("id") &" order by sira asc"
object.open Sorgula,baglanti,1,3  

%>      


<table border="1" bordercolor="#E6E6E6" bgcolor=#FFFFFF cellpadding="3" cellspacing="3" width=700 class=font>
<tr bgcolor=#F3F0E4 height=30>    
   <td><b>Sayfa Adý </b></td>    

      <td align=center><b>Sýralama </b></td>
<td align=center><font color=#0080C0><b>Düzenle</b></td><td align=center><font color=#FF0000><b>SÝL</b></td></tr>      
  
  	 
 
<%
Do while not object.Eof                                                                                                                                                                                                                  
%>   
  

<tr class="menuoff" onmouseover="className='menuon';" onmouseout="className='menuoff';">
 <td><b><%=object("isim")%></td>  

 <td align=center><%=object("sira")%></td>  

<td align=center>
<a href="sayfa-sil.asp?edit=yes&sid=<%=object("id")%>"><img src="images/edit.png" width=20></a>
</td>

<form action="sayfa-sil.asp" method="post" >
<td align=center>    
<input type="hidden" name="id" value="<%=object("id")%>">      
<input type="hidden" name="delete" value="1">        
<input type="image" src="images/del.png"  onclick="return confirm('Sayfayý Silmek Ýstiyor Musunuz?')"></td></form>


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


 