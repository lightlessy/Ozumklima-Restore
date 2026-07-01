   
<!--#INCLUDE file="menu.asp"-->

<body>     

<div class="baslik">Yorumlar </div>

   <%	intGecerliSayfa = Request.Querystring("sayfa")
If intGecerliSayfa = "" OR IsNumeric(intGecerliSayfa) = False Then intGecerliSayfa = 1
%>
   

<BR><BR><center>
                                  




<%if request("delete")<>"" then  %>

<%


Set Rs = Server.CreateObject("Adodb.Recordset")
rSQL = "Select * from yorumlar where id = "& Request.Form("id")
Rs.Open rSQL, baglanti, 1, 3 
Rs.Delete ' Sil
Rs.Update ' Güncelle

rs.close
set rs = nothing 

Response.Redirect Request.ServerVariables("HTTP_REFERER")
%>        
<% end if %>   


<%if request.querystring("update")<>"" then  

Set Rs = Server.CreateObject("Adodb.Recordset")
rSQL = "Select * from yorumlar where id = "& Request("sid")
Rs.Open rSQL, baglanti, 1, 3 

Rs("onay") = 1
Rs.Update 

rs.close
set rs = nothing 

Response.Redirect Request.ServerVariables("HTTP_REFERER")
%>        
<% end if %>

   


<%if request.form("cevapla")<>"" then  

Set Rs = Server.CreateObject("Adodb.Recordset")
rSQL = "Select * from yorumlar where id = "& Request("id")
Rs.Open rSQL, baglanti, 1, 3 

Rs("onay") = 1     
Rs("cevap") =  request.form("cevap") 
Rs.Update 

rs.close
set rs = nothing 

Response.Redirect Request.ServerVariables("HTTP_REFERER")
%>        
<% end if %>       




<% sid = Cint(Request("sid")) %>                         


<BR>


<%  

set object = Server.CreateObJect("ADODB.RecordSet")
Sorgula = "Select * From yorumlar order by id desc"
object.open Sorgula,baglanti,1,3   

if not object.eof then
%>      


<table border="1" bordercolor="#E6E6E6" bgcolor=#FFFFFF cellpadding="3" cellspacing="3" width=980 class=font>
<tr bgcolor=#F3F0E4 height=30>      
<td align=center><b>Yorum</b></td>    
<td align=center><font color=#729A25><b>Cevapla</b></font></td>
<td align=center><font color=#0080C0><b>Onay</b></font></td>
<td align=center><font color=#FF0000><b>SÝL</b></font></td></tr>      


<%
	i = 1 
  	object.Pagesize = 20
	object.absolutepage = intGecerliSayfa
	intSayfaSayisi = object.Pagecount
	
        For b=1 To object.pagesize
	If object.Eof Then Exit For

%>
             
 <form method="post" action="yorumlar.asp"> 
 <tr <% if object("onay")=1 then %> bgcolor=#F3F3F3 <% end if %>> 
 


<td width=444><%=object("yorum")%>
<BR><BR>

<% if request("cevap") = 1 and sid = object("id")  then %>      

<form action="yorumlar.asp" method="post" >
<TEXTAREA   style="WIDTH: 400px; HEIGHT: 90px"  name="cevap"><%=object("cevap") %></TEXTAREA>
<input type="hidden" name="id" value="<%=sid%>">                               
<input type="hidden" name="cevapla" value="1">   
<input type="submit" value="Gönder" ></form>


<% end if %>


</td>  
  
<td><a href="yorumlar.asp?cevap=1&sid=<%=object("id")%>"><font color=#729A25><b>CEVAPLA </b></font> </a></td>

<td align=center width=100>  
<% if object("onay")=1 then %>  <B>Yayýnda</B> 
<% else %>
<a href="yorumlar.asp?update=1&sid=<%=object("id")%>"><font color=#0080C0><b>ONAYLA </b></font> </a>
<% end if %>
</td>
<td align=center><form action="yorumlar.asp" method="post" >
<input type="hidden" name="id" value="<%=object("id")%>">        
<input type="submit" value="SÝL" name="delete" onclick="return confirm('Kayýt Silinecek. Onaylýyor musunuz?')"></td></form>
</div>        



<%
object.MoveNExt
Next
%>        

   </tr></table>  
   <BR><BR>

<table width="900" class=font border=1 bordercolor=#A3D1D1><tr><td align=center>    

<%
Aralik = 2
intSayfaAltSiniri = intGecerliSayfa - Aralik
intSayfaUstSiniri = intGecerliSayfa + Aralik
intSolGrupSayisi = 1 + Aralik
intSagAltSinir = intSayfaSayisi - Aralik
%>


<%
If (intSayfaUstSiniri >intSayfaSayisi) Then
intSayfaAltSiniri = intSayfaSayisi - (2 * Aralik)
intSayfaUstSiniri = intSayfaSayisi
End If
%>


<%
If (intSayfaAltSiniri <= 0) Then
intSayfaAltSiniri = 1
intSayfaUstSiniri = intSayfaAltSiniri + (2 * Aralik)
If intSayfaUstSiniri >= intSayfaSayisi Then intSayfaUstSiniri = intSayfaSayisi
End If
%>



<%
If intSolGrupSayisi >= intSayfaAltSiniri Then intSolGrupSayisi = intSayfaAltSiniri - 1
If intSagAltSinir <= intSayfaUstSiniri Then intSagAltSinir = intSayfaUstSiniri + 1
%>


<%
If intGecerliSayfa >1 Then %>
<a href="?sayfa=1" title="Ýlk Sayfa"><font color=#4E4E4E>[««]</a>
<a href="?sayfa=<%=intGecerliSayfa-1%>" title="Previous Page"><font color=#4E4E4E>[«]</a><%
End If
%>


<%
For i = 1 to intSolGrupSayisi
If CInt(i) = (intGecerliSayfa) Then
Response.Write " <b>"
Else %>
<a href="?sayfa=<%=i%>" title="<%=i%>. Sayfa"><font color=#4E4E4E><%
End If
Response.Write i &".</a></b>"
Next
%>


<%
If (intSayfaAltSiniri - intSolGrupSayisi >1 ) Then Response.Write "<b>. . .</b>"
%>



<%
For i = intSayfaAltSiniri to intSayfaUstSiniri
If CInt(i) = CInt(intGecerliSayfa) Then
Response.Write " <b>"
Else %>
<a href="?sayfa=<%=i%>" title="<%=i%>. Sayfa"><font color=#4E4E4E><%
End If
Response.Write i &".</a></b>"
Next
%>



<%
If (intSagAltSinir - intSayfaUstSiniri >1) Then Response.Write "<b>. . .</b>"
%>





<%
For i = intSagAltSinir to intSayfaSayisi
If CInt(i) = CInt(intGecerliSayfa) Then
Response.Write " <b>"
Else %>
<a href="?sayfa=<%=i%>" title="<%=i%>. Page"><font color=#4E4E4E><%
End If
Response.Write i &".</a></b>"
Next
%>




<%
If CInt(intGecerliSayfa) <>CInt(intSayfaSayisi) Then %>
<a href="?sayfa=<%=intGecerliSayfa+1%>" title="Next Page"><font color=#4E4E4E>[»]</a>
<a href="?sayfa=<%=intSayfaSayisi%>" title="Last Page"><font color=#4E4E4E>[»»]</a><%
End If
%>

</td>	</tr>	</table>  

<%
object.close
set object = nothing
  
end if
%>  

   
             
      
      <BR><BR>


 