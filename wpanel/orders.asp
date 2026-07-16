<html>
<head>      
<meta charset="utf-8">
<link rel="stylesheet" href="print.css"type="text/css" media="print" /> 

<title>Siparişler</title>




   
  <!--#include file="menu.asp"-->      
  
  <div class="baslik">  Devam Eden Siparişler </div>
  
              
               
  <%if request.querystring("del") ="1" then  


Set Rs = Server.CreateObject("Adodb.Recordset")      
rSQL = "Select * from siparis where kod = '"& Request.querystring("kod") & "' "    
Rs.Open rSQL, baglanti, 1, 3   


Set uyem = Server.CreateObject("Adodb.Recordset")
rSQL = "Select * from uyeler where uid =" & Rs("uid") & ""        
uyem.Open rSQL, baglanti, 1, 3   

 sayisi= Rs.recordcount
 
 Dim Dizi()
 ReDim Dizi(sayisi,4) 
  
 k=0                                                    
Do while not Rs.Eof   

set sorad = Server.CreateObJect("ADODB.RecordSet")  
Sorgula = "Select * From products where AffiliateID = "& Rs("product") &"  " 
sorad.open Sorgula,baglanti,1,3   

urunadi=sorad("isim")         
if sorad("indirim")<>0 then
fiyati=FormatNumber(sorad("indirim"),2)  
else
fiyati=FormatNumber(sorad("fiyat"),2)  
end if

miktari=Rs("adet") 
if sorad("indirim")<>0 then
urunt=sorad("indirim")*Rs("adet")  
else
urunt=sorad("fiyat")*Rs("adet")   
end if


tutar=FormatNumber(urunt,2)      
   
Dizi(k,0)=urunadi   
Dizi(k,1)=fiyati
Dizi(k,2)=miktari
Dizi(k,3)=tutar

 
mesaj="<tr class=urunBilgisi><td>"& Dizi(k,0) &" </td><td> "& Dizi(k,1) &" TL </td><td> "& Dizi(k,2) &" Adet </td><td> "& Dizi(k,3) &" TL </td></tr>" 
a= mesaj + a  

k=k+1
Rs.MoveNExt
Loop   




set sepet = Server.CreateObJect("ADODB.RecordSet")
Sorgula = "Select * From siparis where kod = '"& Request.querystring("kod") &"' " 
sepet.open Sorgula,baglanti,1,3   

Do While Not sepet.EOF 

Set stokla = Server.CreateObject("Adodb.Recordset")
rSQL = "Select * from products where AffiliateID = "& sepet("product")
stokla.Open rSQL, baglanti, 1, 3      

stokla("stok") = stokla("stok") - sepet("adet") 
   
stokla.Update 

sepet.MoveNext
Loop      






  
  
  Set Rs = Server.CreateObject("Adodb.Recordset")      
rSQL = "Select * from siparis where kod = '"& Request.querystring("kod") & "' "    
Rs.Open rSQL, baglanti, 1, 3 


Set uyem = Server.CreateObject("ADODB.Recordset")
sql = "SELECT * FROM uyeler WHERE uid =" & Rs("uid") & " "
uyem.open sql, baglanti,1,3    

 
	
Session.CodePage = 65001                      
kime=""&uyem("mail")&""
kimden="form@keramikadenizli.com" 
konu="Siparisiniz Kargoya Verildi"    
isim="Keramika Denizli"
mesajm="<table cellspacing=4 cellpadding=4 bgcolor=#EFEFEF width=800 style=""font-family:Open Sans; font-size:15px;"" border=0><tr><td><a href=http://keramikadenizli.com target=_blank><img src=http://keramikadenizli.com/images/logo.png border=0><BR><BR></td><td align=center><a href=http://keramikadenizli.com target=_blank style=""text-decoration: none""><font color=#619106 size=4><B>www.keramikadenizli.com</B></font></a></td></tr><tr><td bgcolor=#3E3E3E  colspan=2><font   color=#FFFFFF>Merhaba, Siparişiniz Kargolandı.</font></td></tr><tr><td  colspan=2>Kargo Takip Barkod Numarası: "& Rs("takip") &"  <BR> <a href=http://gonderitakip.ptt.gov.tr target=_blank><B>Kargo Takibi İçin Tıklayın</B> </a></td></tr><tr><td colspan=2>Siparişinizdeki ürünlerin tamamı kargoya verilmiştir. Kargonuzun durumunu web sitemizdeki siparişlerim bölümünden takip edebilirsiniz. </td></tr><tr><td colspan=2><table width=750 cellspacing=0 cellpadding=0 border=1 bordercolor=#CACACA style=""padding:7px; font-size:13px"">"& mesaj &"</table></td></tr><tr><td bgcolor=#619106 colspan=2 height=35><a href=http://keramikadenizli.com target=_blank style=""text-decoration: none""><font color=#FFFFFF><B>Mutfağınızın, sofralarınızın ve evinizin rengi... » www.keramikadenizli.com</B></font></a></td></tr><tr><td style=""font-size:12px""> Tel : "& ayars("tel") &"   <BR>Adres: DENİZLİ <BR> E-Posta: <a href=mailto:"& ayars("eposta") &" target=_blank>"& ayars("eposta") &"</a>       <BR>Sosyal Medya: <BR> <a href="& ayars("face") &" target=_Blank>Facebook</a>  <a href="& ayars("twit") &" target=_Blank>Twitter</a> <a href="& ayars("insta") &" target=_Blank>Instagram</a></td><td align=right><img src=http://keramikadenizli.com/images/logo.png border=0 width=150></td></tr></table>"    

	Set mektup = Server.Createobject("Persits.MailSender")
	mektup.Host = "mail.keramikadenizli.com"
	mektup.IsHTML =true 
	mektup.From = kimden
	mektup.FromName = isim
	mektup.AddAddress kime
	mektup.AddReplyTo kimden
	mektup.Subject =konu
	mektup.Body = mesajm
	On Error Resume Next    
	mektup.Username = "form@keramikadenizli.com"
        mektup.Password = "SGii53T8"    
        mektup.SSL = True
	mektup.Port = 465    
        mektup.CharSet = "UTF-8"
        mektup.ContentTransferEncoding = "Quoted-Printable"
	mektup.Send
	Set mektup = Nothing   
	
	
	
	
	  On Error Resume Next
Function HTTPPoster(ByVal prmPostAddress, ByVal prmSendData)
	Set objXMLHTTP = Server.CreateObject("Microsoft.XMLHTTP")
	objXMLHTTP.open "POST",prmPostAddress,False
	objXMLHTTP.send prmSendData
	if Err then
		Err.Clear
		HTTPPoster = ""
		Set objXMLHTTP = Nothing
		Exit Function
	else
		HTTPPoster = objXMLHTTP.responseText
		Set objXMLHTTP = Nothing
	end if
End Function

Sub postsms()
	StrXML = "<SMS>"&_
	   "<oturum>"&_
		  "<kullanici>8506761031</kullanici>"&_
		  "<sifre>dz8gk4</sifre>"&_
	   "</oturum>"&_
	   "<mesaj>"&_
		  "<baslik>KERAMIKADNZ</baslik>"&_
		  "<metin>Siparişinizdeki ürünlerin tamamı kargoya verilmiştir. Kargonuzun durumunu web sitemizdeki siparişlerim bölümünden takip edebilirsiniz. </metin>"&_
		  "<alicilar>["&uyem("tel")&"]</alicilar>"&_
	   "</mesaj>"&_
	"</SMS>"
	Response.write HTTPPoster("http://dakikapi.com/api/xml_api.php", strXML)
End Sub
postsms() 	
     

Do While Not Rs.EOF    
Rs.Delete       
Rs.Update 
Rs.MoveNext
Loop    


rs.close
set rs = nothing        

Set upd = Server.CreateObject("Adodb.Recordset")      
rSQL = "Select * from arsiparis where kod =  '"& Request.querystring("kod") & "' "    
upd.Open rSQL, baglanti, 1, 3 

upd("onay") = 1  
upd.Update           



Response.Redirect "orders.asp" 
%>        
<% end if %>       


</HEAD>   
<body>



  <BR><BR>     


 <%         
 if request.querystring("detail")=1 then   
 
 
set mods = Server.CreateObJect("ADODB.RecordSet")
Sorgula = "Select * From siparis where kod = '"& request.querystring("ord") &"' " 
mods.open Sorgula,baglanti,1,3 %>
                                                                   
                                          


<table bgcolor=#F2F2F2 cellpadding=8 cellspacing=8 align=center class=font border=0 align=center>   
     
 <tr> <td> Sipariş Kodu: <B><%=mods("kod")%></B> <BR> <%=mods("tarih")%>      </td>
 <td align=center> <% if mods("takip") <>"" then %><a href="orders.asp?del=1&kod=<%=mods("kod")%>" onclick="return confirm('Sipariş Kargolandı. Onaylıyor Musunuz?')">
 <div style="background-color:#008080; padding:8px"><font color="#FFFFFF" size=3>Siparişi Tamamla</font>  </div></a>
 <% else %>      
  <div style="background-color:#6A6A6A; padding:8px"><font color="#FFFFFF" size=2> Siparişi Tamamlamak İçin Kargo Takip No Giriniz</font>  </div>
 <% end if %>         </td></tr>   

      

<tr><td colspan=2>        
<%
Set Rsp = Server.CreateObject("Adodb.Recordset")
rSQL = "Select * from uyeler where uid =" & mods("uid") & ""        
Rsp.Open rSQL, baglanti, 1, 3  %>      

                                                     
<% If Rsp.EOF And Rsp.BOF Then %>     
                      
<font class=fy>
<font color=#CE0000><B>Müşteriye Ulaşılamadı !</B>

<% else %>

<div style="padding-top:2px; padding-left:3px; padding-right:6px; padding-bottom:0px; line-height:18px ">
       


<table  class=font border=0 width=700 cellpadding=2 cellspacing=2 bgcolor=#FFFFFF bordercolor=#ACACAC >          
<tr><td height=35 colspan=2 align=center>
<B>            
<% if mods("odeme")=0 then %> <div style="background-color:#939393; color:#FFFFFF; font-size:20px;  padding:15px; text-align:center;"> ÖDEME İLE İLGİLİ HERHANGİ BİR İŞLEMDE BULUNULMADI  </div><% end if %>
<% if mods("odeme")=1 then %>  <div style="background-color:#008ED2; color:#FFFFFF; font-size:20px;  padding:15px; text-align:center;"> KREDİ KARTI İLE ÖDEME ALINDI    </div><% end if %>
<% if mods("odeme")=2 then %> <div style="background-color:#376088; color:#FFFFFF; font-size:20px;  padding:15px; text-align:center;"> BANKA HAVALESİ / EFT ALINDI    </div><% end if %>
<% if mods("odeme")=3 then %><div style="background-color:#376088; color:#FFFFFF; font-size:20px;  padding:15px; text-align:center;">  KAPIDA ÖDEME TALEBİ ALINDI   </div> <% end if %>
</B>

</td></tr> 
<tr bgcolor=#F3F3F3><td height=35 colspan=2> &nbsp;&nbsp; <B>TESLİMAT ADRESİ</B><BR></td></tr> 
<tr><td height=25>  &nbsp;&nbsp; Ad Soyad:</td><td ><%=Rsp("ad")%>&nbsp;<%=Rsp("soyad")%></td></tr>
<tr><td height=25>       &nbsp;&nbsp; Adres:  </td><td> <%=Rsp("adres")%> / <%=Rsp("sehir")%></td></tr>      
<tr><td height=25>  &nbsp;&nbsp; Tel:  </td><td><%=Rsp("tel")%></td></tr>   
<tr><td height=25>  &nbsp;&nbsp; TC No:  </td><td><%=Rsp("tcno")%></td></tr>   
<tr><td height=25>  &nbsp;&nbsp; e-Mail:  </td><td><%=Rsp("mail")%></td></tr>  
</table>  



      <%if request("1")<>"" then 


Set Rs = Server.CreateObject("Adodb.Recordset")
rSQL = "Select * from siparis where kod = '"& Request.Form("id") &"' "
Rs.Open rSQL, baglanti, 1, 3    

Do While Not Rs.EOF 
  
Rs("takip") = Request("takip")    
Rs.Update    

Rs.MoveNext
Loop      

rs.close
set rs = nothing  


Set Rs = Server.CreateObject("Adodb.Recordset")
rSQL = "Select * from arsiparis where kod = '"& Request.Form("id") &"' "
Rs.Open rSQL, baglanti, 1, 3    

Do While Not Rs.EOF 
Rs("takip") = Request("takip")    
Rs.Update    

Rs.MoveNext
Loop      

rs.close
set rs = nothing 

Response.Redirect Request.ServerVariables("HTTP_REFERER")
%>        
<% end if %>

 
    

</div>       

<% end if %>   

</td></tr>      

<tr><td colspan=2>    
  
    <form method="post" action="orders.asp?detail=1&ord=<%=request.querystring("ord") %>#k"  > 
 
<div style="background-color:#0080C0; color:#FFFFFF; padding:4px; ">
Kargo Takip No: <input type="text" size="10" name="takip" value="<%=mods("takip")%>">  
<input type="hidden"  name="id" value="<%=request.querystring("ord") %>">            
 <input type="submit" name="1" value="Güncelle" />
 </div>
  </form>      
  
                    
<table bgcolor=#FFFFFF class=font  border=0 cellspacing=4 cellpadding=4>  
<tr><td><b>Sipariş Listesi: </b></td></tr>             
 <tr bgcolor=#F3F3F3>
<td height=19 width=485><B>Ürünler:</B></td>  
<td width=55><B>Fiyat:</B></td>    
<td width=55><B>Miktar:</B></td>  
<td align=right><B>Toplam:</B></td>
</tr>     

           

<%Do while not mods.Eof  

set sorad = Server.CreateObJect("ADODB.RecordSet")
Sorgula = "Select * From products where AffiliateID = "& mods("product") &"  "
sorad.open Sorgula,baglanti,1,3   

if sorad.eof then %>
<BR><BR><font color=#A40000><B>Ürüne Ulaşılamıyor</B> </font><BR><BR>
<% else


set sorface = Server.CreateObJect("ADODB.RecordSet")
Sorgula = "Select * From surface where grup = "& mods("product") &"  "
sorface.open Sorgula,baglanti,1,3       


%>                          
                                                                                                                                      
<tr>
<td height=20 valign=top> <a href="http://www.keramikadenizli.com/<%=sorad("isim")%>-<%=sorad("AffiliateID")%>" target=_blank style="color:#454545"><B><%=sorad("isim")%></B> </a>
</td>
<td valign=top align=center><% if sorad("indirim")=0 then %> <%=FormatNumber(sorad("fiyat"),2)%> <% else %>  <%=FormatNumber(sorad("indirim"),2)%>  <% end if %>  TL</td>     
<td valign=top align=center>    <%=mods("adet") %> </td>
<td valign=top align=right> 

<% if sorad("indirim")=0 then %> 
<%urunt=sorad("fiyat")*mods("adet")%> <%=FormatNumber(urunt,2)%>
 <% else %>  
<%urunt=sorad("indirim")*mods("adet")%> <%=FormatNumber(urunt,2)%>  <% end if %>  TL </td> </tr>   


                 
<% end if %>           


<%  
mods.MoveNExt
Loop


set modss = Server.CreateObJect("ADODB.RecordSet")
Sorgula = "Select * From siparis where kod = '"& request.querystring("ord") &"' " 
modss.open Sorgula,baglanti,1,3 
%> 
<tr  bgcolor=#F2F2F2 height=30><td colspan=5  align=right>   Kargo:   <% if modss("kargo")=0 then %><font color="#FF0080">BEDAVA  </font> <% else %>    <%=FormatNumber(modss("kargo"),2)%> TL  &nbsp;  <% end if %></td></tr> 

<tr  bgcolor=#FFFFFF height=30><td colspan=5  align=right>   Toplam:    <B>   <%=FormatNumber(modss("tfiyat"),2)%> TL  </B>&nbsp;</td></tr> 
       
</table> 
<BR>      
 <a href="orders.asp"><font color=#BF0000 size=2> «  Tüm Siparişler </font></a>            
  
        </td></tr> 
       
</table>     
  
  
    <BR><BR>&nbsp;
  
  
  
  
  
  
  
  
  
  
   

<% else %>



<TABLE border=0 class=font cellspacing=6 cellpadding=6 align=center>
             
<%  
set mods1 = Server.CreateObJect("ADODB.RecordSet")
Sorgula = "Select kod, uid, tarih, onay From siparis where onay=1 group by  kod,uid, tarih, onay order by tarih desc"
mods1.open Sorgula,baglanti,1,3      
   
%>   
           
                                                     
<% If mods1.EOF And mods1.BOF Then %>          

<TR><TD valign=top>   
                                <BR><BR>
<font class=fy><font color=#CE0000><B>Sipariş Yok </B> </font></font>        
 
</TD></TR>

<% else %>



   <div class=print>  

<TD valign=top>        
                              

<% Do while not mods1.Eof  %>        


<table bgcolor=#FFFFFF width=500 cellpadding=5 cellspacing=5 align=center class=font border=0>   
     <tr> <td>  Sipariş Kodu: <B><%=mods1("kod")%></B> - <%=mods1("tarih")%>                   </td>
<td align=center> <a href="order-sil.asp?sil=<%=mods1("kod")%>" onclick="return confirm('Kayıt silinecek. Onaylıyor Musunuz?')"><img src="images/del.png" width=20></a>
</td></tr>   
               
<tr><td height=20>          <a href="orders.asp?detail=1&ord=<%=mods1("kod")%>"><font color="#0080C0"><B>Sipariş Dökümü »</B></font> </a>     </td></tr>                 
</table> 
 <BR>
 
    <%  
mods1.MoveNExt
Loop
%>     

   
             
</div>


</TD></TR></TABLE>   


<% end if %>
     
<% end if %> 


      </body>
      </html>
      