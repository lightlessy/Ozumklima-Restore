
  <!--#include file="menu.asp"-->  
 

  <div class="baslik">SMS Gönder</div>    
                                               
<% tip= request.querystring("tip") %>

  <body OnLoad="initialiseWebWizRTE();">  


<% 
Function KarakterTemizle(Veri)
KarakterTemizle = Replace(Veri, "'", "''")
Veri = KarakterTemizle                                                         
End Function            %>

 
 <% if  request("1") <>""  then  


mesaj = request.form("mesaj") 
 
 Dim alici  

Set Rs = Server.CreateObject("Adodb.Recordset")
rSQL = "Select * from uyeler "
Rs.Open rSQL, baglanti, 1, 3          

 
Do while not Rs.Eof   

a = "["& Rs("tel") &"]," + a 
        
Rs.MoveNExt 
Loop 



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
		  "<metin>"& mesaj &"</metin>"&_
		  "<alicilar>"& a &"</alicilar>"&_
	   "</mesaj>"&_
	"</SMS>"
	Response.write HTTPPoster("http://dakikapi.com/api/xml_api.php", strXML)
End Sub
postsms()   

%>                          
          
<div class=font style="margin:0px auto; width:500px; text-align:center; margin-top:150px; font-size:1.2em; color:#0080C0">
<B>SMS GÖNDERÝLDÝ</B>
<%=a%>
</div>

<%else%>   
          
    
    <BR><BR>
         
  <form method="post" action="smsat.asp" name="myForm" id="myForm">         
   	   
 <table border=1 bgcolor=#FFFFFF bordercolor=#E6EBEE cellpadding=5 cellspacing=5 align=center  class=font align=center>  
 

<tr><td>    
  <b> Alýcý </b>:   Tüm Üyeler</td></tr>       

<tr><td>    
  <b> Mesaj Ýçeriði </b>:
  <BR><BR>
   
  <TEXTAREA cols=100 name="mesaj"  rows=5 required></TEXTAREA></td></tr>      
   
</tr>	     
   
  <tr><td colspan=4 align=center>
  <input type="submit" name="1" value="GÖNDER" />
</td></tr>
</table>
      </form>       
                               
      
<%   end if  %>
