<html>
 <head>
  <title>PANEL </title> 
  <meta http-equiv="Content-Type" content="text/html; charset=utf-8">        
<link rel=stylesheet type="text/css" href="font.css">
</head>     
   
   
  <!--#include file="menu.asp"-->  
  
<div class="baslik">Haber Beklenen Ürünler</div>
<BR><BR>

<%
if Request.Querystring("sil")=1 then 


Set Rs = Server.CreateObject("Adodb.Recordset")
rSQL = "Select * from haberver  where grup = "& Request("seri")
Rs.Open rSQL, baglanti, 1, 3 
     
Do while not Rs.Eof                                                                                                                                                                                                                  
 
 Rs.Delete 
Rs.Update  
  
Rs.MoveNExt
Loop


rs.close
set rs = nothing 
Response.Redirect Request.ServerVariables("HTTP_REFERER")         

end if      %>   

                
 
 
 <%
if Request.Querystring("sms")=1 then 

Dim alici  

Set Rs = Server.CreateObject("Adodb.Recordset")
rSQL = "Select * from haberver  where grup = "& Request("seri")  &""
Rs.Open rSQL, baglanti, 1, 3          

 
Do while not Rs.Eof   

a = "["& Rs("cep") &"]," + a 
        
Rs.MoveNExt 
Loop   


Set mods = Server.CreateObject("Adodb.Recordset")
rSQL = "Select * from products where AffiliateID = "& Cint(Request("seri")) &""
mods.Open rSQL, baglanti, 1, 3   


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
		  "<metin>Merhaba. keramikadenizli.com web sitemizde beklediğiniz "& mods("isim") &" ürünü stoklarımızda!</metin>"&_
		  "<alicilar>"& a &"</alicilar>"&_
	   "</mesaj>"&_
	"</SMS>"
	Response.write HTTPPoster("http://dakikapi.com/api/xml_api.php", strXML)
End Sub
postsms() 


Set Rs2 = Server.CreateObject("Adodb.Recordset")
rSQL = "Select * from haberver  where grup = "& Request("seri")  &""
Rs2.Open rSQL, baglanti, 1, 3  

Do while not Rs2.Eof   

Rs2("gitti") =1
Rs2.Update   

Rs2.MoveNExt 
Loop   

   
end if      %>  



<center><font class=fontbig>
<B>Haber Beklenen Ürünler </B><BR><BR>
     
</font>
                            

<table class=font border=1 bordercolor=#E8E8E8 cellspacing=3 cellpadding=4 width=780 align=center>

 
<%

Set mods1 = Server.CreateObject("Adodb.Recordset")
rSQL = "Select grup,isim, gitti  from  haberver group by grup, isim, gitti "
mods1.Open rSQL, baglanti, 1, 3     



Do while not mods1.Eof     

Set mods = Server.CreateObject("Adodb.Recordset")
rSQL = "Select * from haberver  where grup="& mods1("grup") &""
mods.Open rSQL, baglanti, 1, 3       

%>  
           
<tr bgcolor=#F3F0E4><td ><B><%=mods1("isim")%></B>  <a href="urun-beklenen.asp?seri=<%=mods1("grup")%>&sil=1" onclick="return confirm('Kayıtlar Silinecek. Onaylıyor Musunuz?')"><div style="float:right; color:#D70000"> <B>Kayıtları Sil</B></div> </a></td></tr> 
    
    
<% 
  k=1 
  Do while not mods.Eof      %>   
  

<tr class="menuoff" onmouseover="className='menuon';" onmouseout="className='menuoff';"><td><%=mods("cep")%></td></tr>
<%   
  k = k + 1 
mods.MoveNExt
Loop
%>      
</td></tr>   

<tr bgcolor="#3F3F3F"><td><% if mods1("gitti")=1 then %> <font color=#FFFFFF>SMS Gönderildi </font><% else %> <a href="urun-beklenen.asp?seri=<%=mods1("grup")%>&sms=1" style="color:#FBBF44"><B>ÜRÜN GELDİ SMS GÖNDER</B> </a> <% end if %> </td></tr>  

<tr><td><BR></td></tr>  



<%     

mods1.MoveNExt
Loop
%>  
</table>  

                         
<BR><BR>
&nbsp;

