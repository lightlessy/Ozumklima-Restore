<html>
<head>    
<title>Müşteri Hesabı Düzenle</title>
  <meta http-equiv="Content-Type" content="text/html; charset=windows-1254">        
<link rel=stylesheet type="text/css" href="font.css">
</head>   

 
<!--#INCLUDE file="dbase.asp"-->   

 
   

</HEAD>               

 
 <% if request("update")<>"" then  %>     
 
 <%
tel1= "" & Request.Form("tela")  & "" & Request.Form("cep")  & ""     
    

Set Rs = Server.CreateObject("Adodb.Recordset")
rSQL = "Select * from uyeler where uid =" & request.form("uid") & ""
Rs.Open rSQL, baglanti, 1, 3                  

Rs("ad") = Request("ad")
Rs("soyad") = Request("soyad")      
Rs("kullanici") = Request("kullanici")  
Rs("sifre") = Request("sifre")  
Rs("cep") = Request("cep")  
Rs("tel") = Request("tel")  
Rs("tarih") = Request("tarih")  
Rs("adres") = Request("adres")   
Rs("mail") = Request("mail")  
Rs.Update    

rs.close
set rs = nothing    
Response.Redirect Request.ServerVariables("HTTP_REFERER")
          end if
  %>

       
<body>  

<%
Set rs = Server.CreateObject("ADODB.Recordset")
sql = "SELECT * FROM uyeler WHERE uid = "& request("id")
rs.open sql, baglanti,1,3    
%>     

               <BR>
<TABLE class=fontm width=470 border=0><TR><TD valign=top>

<table bgcolor=#F5F5F5 bordercolor=#EAEAEA cellpadding=5 cellspacing=5 align=center class=font border=1> 

<form method="POST" action="uye-edit.asp?id=request("id")" name="myForm" id="searchform" onSubmit="return CheckMSForm();">       
<input type="hidden" size="20" name="ad" value="<%=rs("ad")%>">
<input type="hidden" size="20" name="soyad" value="<%=rs("soyad")%>">

  
  <tr><td><b>Ad Soyad</b></td><td><%=rs("ad")%>&nbsp;<%=rs("soyad")%></td></tr>

  <tr><td><b>Kullanıcı</b></td><td><input type="text" size="30" name="kullanici" value="<%=rs("kullanici")%>"></td></tr>
<tr><td><b>Şifre</b> </td><td> <input type="text" size="20"  name="sifre" value="<%=rs("sifre")%>"></td></tr>     
    
<tr>
<td> <b>Adres</b></td><td><TEXTAREA  title="Lütfen adresinizi eksiksiz yazınız" style="WIDTH: 270px; HEIGHT: 55px"  name=adres><%=rs("adres")%></TEXTAREA></td></tr>

<tr><td><b>Tel</b></td><td>
 <input type="text" size="14" name="tel"  value="<%=rs("tel")%>" onkeypress="return numbersonly(this, event)">
  </td>    </tr>
  

 <tr> <td><b>E-mail</b></td><td><input type="text" size="30" name="mail" value="<%=rs("mail")%>"></td></tr>     


<input type="hidden" name="tarih"   size="12" value="<%=date%>">
<input type="hidden" name="uid"   size="12" value="<%=rs("uid")%>">
 
<tr><td colspan=4 align=center>
<input type="submit" name="update" value="Güncelle"></td>

</form></tr>
</table>
         


           
      </body>
      </html>