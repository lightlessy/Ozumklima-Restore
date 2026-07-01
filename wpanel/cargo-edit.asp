
  <!--#include file="menu.asp"-->    
  
     <div class="baslik">  Kargo Düzenleme</div>
  
  <body>        
  


 <% if  request.form("kargo") <>""  then  


set Rs = Server.CreateObJect("ADODB.RecordSet")
Sorgula = "Select * From cargo"
Rs.open Sorgula,baglanti,1,3    

           
Rs("kargo")=   Request.Form("kargo")      
Rs("bedava")=   Request.Form("bedava") 
Rs.Update            

Response.Redirect("cargo-edit.asp")

else %>       

<BR><BR>

 <%  

set object = Server.CreateObJect("ADODB.RecordSet")
Sorgula = "Select * From cargo"
object.open Sorgula,baglanti,1,3  
%>   
     
 <table border=1 bordercolor=#E6EBEE bgcolor=#FFFFFF cellpadding=5 cellspacing=5 align=center  class=font>                       

<FORM  name="f"  action="cargo-edit.asp" method=post  >                    

<tr><td>Genel Kargo Fiyatý</td><td>     <INPUT   style="WIDTH: 50px"  name="kargo" value="<%=object("kargo")%>">  TL     </td></tr> 

<tr><td>Bedava Limiti</td><td>     <INPUT   style="WIDTH: 50px"  name="bedava" value="<%=object("bedava")%>">  TL     </td></tr>

<tr> <td colspan=2 align=center><input type="submit" value="Güncelle">   </td></tr>

 </FORM>
 </table>
               
 <BR><BR>&nbsp;
       
 <% end if %>  
 
