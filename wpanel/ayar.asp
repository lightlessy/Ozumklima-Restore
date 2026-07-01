   
<!--#INCLUDE file="menu.asp"-->  

 <div class="baslik">Ayarlar</div>  

  <body>  
  


<BR><BR><center>
                                                  

<%if request("update")<>"" then 


Set Rs = Server.CreateObject("Adodb.Recordset")
rSQL = "Select * from ayarlar "
Rs.Open rSQL, baglanti, 1, 3 

Rs("adres") = Request("adres")
Rs("tel") = Request("tel")    
Rs("faks") = Request("faks")    
Rs("eposta") = Request("eposta")
Rs("gsm") = Request("gsm")   
Rs("gsm2") = Request("gsm2")     
Rs("kullanici") = Request("kullanici")
Rs("sifre") = Request("sifre") 
Rs("banka1") = Request("banka1")
Rs("banka2") = Request("banka2")
Rs("banka3") = Request("banka3")
Rs("banka4") = Request("banka4")    
Rs("face") = Request("face")    
Rs("twit") = Request("twit")    
Rs("insta") = Request("insta")    
Rs("you") = Request("you")    
Rs("formail") = Request("formail")    
Rs("firma") = Request("firma")      
Rs("map") = Request("map")   
Rs.Update 

rs.close
set rs = nothing 

Response.Redirect Request.ServerVariables("HTTP_REFERER")
%>        
<% end if %>


<BR>


      <%    

set object = Server.CreateObJect("ADODB.RecordSet")
Sorgula = "Select * From ayarlar"
object.open Sorgula,baglanti,1,3             

%>      


 <form method="post" name="myForm" action="ayar.asp"> 
   	 
<table border="1" bordercolor="#E6E6E6" bgcolor=#FFFFFF cellpadding="3" cellspacing="3" width=800 class=font>          

<tr><td><b>Firma Adý:</b> </td><td><input type="text" name="firma"   value="<%=object("firma")%>" size="55"></td></tr>   

<tr><td colspan=2 bgcolor="#4D4D4D" style="color:#FFFFFF"> Ýletiþim Bilgileri </td></tr>  

<tr bgcolor="#F0F0E8"><td><b>Adres:</b> </td><td> <TEXTAREA cols=50 name="adres"  rows=5><%=object("adres")%></TEXTAREA></td></tr>   
<tr bgcolor="#F0F0E8"><td><b>Google Map:</b> </td><td> <TEXTAREA cols=50 name="map"  rows=5><%=object("map")%></TEXTAREA></td></tr>  

<tr bgcolor="#F0F0E8"><td><b>Telefon:</b> </td><td><input type="text" name="tel"   value="<%=object("tel")%>" size="55"></td></tr>  
<tr bgcolor="#F0F0E8"><td><b>Faks:</b> </td><td><input type="text" name="faks"   value="<%=object("faks")%>" size="55"></td></tr>  
<tr bgcolor="#F0F0E8"><td><b>GSM 1:</b> </td><td><input type="text" name="gsm"   value="<%=object("gsm")%>" size="55"></td></tr>  
<tr bgcolor="#F0F0E8"><td><b>GSM 2:</b> </td><td><input type="text" name="gsm2"   value="<%=object("gsm2")%>" size="55"></td></tr>     
<tr bgcolor="#F0F0E8"><td><b>E-Mail:</b> </td><td><input type="text" name="eposta"   value="<%=object("eposta")%>" size="55"></td></tr>  
<tr bgcolor="#F0F0E8"><td><b>Ýletiþim Formu Maillerini Almak Ýstediðiniz Mail Adresi:</b> </td><td><input type="text" name="formail"   value="<%=object("formail")%>" size="55"></td></tr>           

<tr><td colspan=2 bgcolor="#4D4D4D" style="color:#FFFFFF"> Giriþ Bilgileri </td></tr>     

<tr bgcolor="#E7EEF1"><td><b>Kullanýcý Adý:</b> </td><td><input type="text" name="kullanici"   value="<%=object("kullanici")%>" size="55"></td></tr>  
<tr bgcolor="#E7EEF1"><td><b>Þifre:</b> </td><td><input type="text" name="sifre"   value="<%=object("sifre")%>" size="55"></td></tr>  

 
<tr><td colspan=2 bgcolor="#4D4D4D" style="color:#FFFFFF"> Banka Hesap Bilgileri </td></tr>     

<tr bgcolor="#EFF2E6"><td><b>Banka 1</b> </td><td><TEXTAREA cols=50 name="banka1"  rows=5><%=object("banka1")%></TEXTAREA></td></tr>  
<tr bgcolor="#EFF2E6"><td><b>Banka 2</b> </td><td><TEXTAREA cols=50 name="banka2"  rows=5><%=object("banka2")%></TEXTAREA></td></tr>  
<tr bgcolor="#EFF2E6"><td><b>Banka 3</b> </td><td><TEXTAREA cols=50 name="banka3"  rows=5><%=object("banka3")%></TEXTAREA></td></tr>  
<tr bgcolor="#EFF2E6"><td><b>Banka 4</b> </td><td><TEXTAREA cols=50 name="banka4"  rows=5><%=object("banka4")%></TEXTAREA></td></tr>      

<tr><td colspan=2 bgcolor="#4D4D4D" style="color:#FFFFFF"> Sosyal Medya Linkleri  </td></tr>     
<tr bgcolor="#E4E1F7"><td><b>Facebook:</b> </td><td><input type="text" name="face"   value="<%=object("face")%>" size="55"></td></tr>
<tr bgcolor="#E4E1F7"><td><b>Twitter:</b> </td><td><input type="text" name="twit"   value="<%=object("twit")%>" size="55"></td></tr>
<tr bgcolor="#E4E1F7"><td><b>Instagram:</b> </td><td><input type="text" name="insta"   value="<%=object("insta")%>" size="55"></td></tr>
<tr bgcolor="#E4E1F7"><td><b>You Tube:</b> </td><td><input type="text" name="you"   value="<%=object("you")%>" size="55"></td></tr>

<tr><td align=center colspan=3>
<input type="submit" name="update" value="Güncelle">

</td>      </form>


   </tr></table>  
   
 





 