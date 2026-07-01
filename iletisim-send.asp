               <CENTER>        
               <span style="font-family:arial; font-size:1em">
<META HTTP-EQUIV="Content-Type" content="text/html; charset=iso-8859-9">   


 <BR><BR><BR> <BR><BR>
 
 
<%    

 
isim=request.form ("isim") 
mail=request.form ("mail")   
tel=request.form ("tel")      
mesaj=request.form("mes")        



Session.CodePage = 1254              
mesaj="<table style=""font-family:arial; font-size:14px; width:300px; border:1px solid #E6E6E6""  cellspacing=6 cellpadding=5 ><tr><td colspan=2><font color=#0080C0><B>Bize Ulaşın Formu</B></font></td></tr><tr><td>Ad Soyad:</td><td><b>" & isim & "</b></td></tr><tr><td>E-Posta:</td><td><b>" & mail & "</b></td></tr><tr><td>Telefon:</td><td><b>" & tel & "</b></td></tr><tr bgcolor=#F3F3F3><td>Mesaj:</td><td> <b>" & mesaj & "</b></td></tr></table>"          
kime="info@ozumklima.com"
kimden="form@ozumklima.com" 
konu="Bize Ulasin"     
isim="Iletisim Formu"
	Set mektup = Server.Createobject("Persits.MailSender")
	mektup.Host = "mail.ozumklima.com"
	mektup.IsHTML =true 
	mektup.From = kimden
	mektup.FromName = isim
	mektup.AddAddress kime
	mektup.AddReplyTo kimden
	mektup.Subject =konu
	mektup.Body = mesaj
	'On Error Resume Next    
mektup.Username = "form@ozumklima.com"  
mektup.Password = "RFit30X4"        
mektup.CharSet = "UTF-8"     
mektup.Port = 587  
        mektup.ContentTransferEncoding = "Quoted-Printable"
	mektup.Send
	Set mektup = Nothing    
	
	%>

<span style="font-family:arial; font-size:1em">	     

<B>Mesajınız iletilmiştir. Teşekkürler.</B>

<BR><BR>

<meta http-equiv="refresh"
content="1; URL=https://ozumklima.com">

</span>
	 
	

	
                    
   
   
   
   
