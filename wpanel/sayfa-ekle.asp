
  <!--#include file="menu.asp"-->  
 

  <div class="baslik">Ýçerik Düzenle</div>    
                                               
<% tip= request.querystring("tip") %>

  <body OnLoad="initialiseWebWizRTE();">  


<% 
Function KarakterTemizle(Veri)
KarakterTemizle = Replace(Veri, "'", "''")
Veri = KarakterTemizle                                                         
End Function            %>

 
 <% if  request("1") <>""  then  


set Rs = Server.CreateObJect("ADODB.RecordSet")
Sorgula = "Select * From sayfalar"
Rs.open Sorgula,baglanti,1,3    

Rs.AddNew            
Rs("icerik")=   KarakterTemizle(Request.Form("icerik"))
Rs("sira")=   Request.Form("sira") 
Rs("isim")=   KarakterTemizle(Request.Form("isim"))    
Rs("yayin")=   Request.Form("yayin")    
Rs("ana")=   Request.Form("ana") 
Rs.Update   
Response.redirect("sayfa-sil.asp?id="&Request.Form("ana")&"")  %>                          


<%else%>   
          
    
    <BR><BR>
         
  <form method="post" action="sayfa-ekle.asp" name="myForm" id="myForm">         
   	   
 <table border=1 bgcolor=#FFFFFF bordercolor=#E6EBEE cellpadding=5 cellspacing=5 align=center  class=font align=center>  
 
  <tr><td><b>Sayfa  Kategorisi </b></td><td  colspan=3>

<select name="ana"> 
  
<option value="0">Lütfen Seçiniz</option>   
<option  value="1">Kurumsal</option>    
<option  value="2">Hizmet</option> 
<option  value="3">Blog</option>     
</select>  

</td></tr>	     
 
  

<tr><td width=240><font color=#000000><b>Sayfa Adý </td><td  colspan=3><input type="text" size="100" name="isim"></td></tr>	

<tr><td> <font color=#000000>  <b>Ýçerik </b>:</td><td colspan=3>     

	<%
strFormName = "myForm"
strTextAreaName = "icerik"
%> 
     
    <!--#include file="RTE_editor_inc.asp" --><BR>
   <TEXTAREA cols=100 name="icerik" id="icerik" rows=25></TEXTAREA></td></tr>      
   
 <tr><td><B>Yayýn:</B> </td><td><input type="checkbox" name="yayin"  value="1"></td></tr>           
   
 <tr><td> <b>Sayfa Sýralama </b> </td><td><input type="text" size="3" name="sira" onkeypress="return numbersonly(this, event)" value="0"></td>	  

</tr>	     
   
  <tr><td colspan=4 align=center>
  <input type="submit" name="1" value="GÖNDER" />
</td></tr>
</table>
      </form>       
                               
      
<%   end if  %>
