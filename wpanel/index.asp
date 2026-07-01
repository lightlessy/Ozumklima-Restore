<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1254"> 
<title> WEB PANEL | Giriş</title>
</head>



<%lg = request("lg")%>

  

<body background="images/arka.jpg">        

<link rel=stylesheet type="text/css" href="font.css">      


 <div style="padding-top:56px; margin:0px auto; width:700px; text-align:center" class="font">    
 

<form method="post" name="f1" action="login2.asp">


<table class="font" border=0 cellspacing=0 cellpadding=0 align=center  background="images/log.jpg" width=375 height=231 align=center>
 <tr><td align=center>
<div style="width:350px; margin:0px auto; padding-top:70px; color:#353535; font-weight:bold">  

 <%if lg="nousername" then%> 
  <BR><font class=fontmain><font color=red><b>Hatalı Kullanıcı Bilgileri</b></font> 
<%end if%> 

<div style="display:inline-block; width:95px; margin-bottom:15px; text-align:left;">Kullanıcı Adı: </div> <div style="display:inline-block;"> <input type="text" size="15" name="username" maxlength="15" style="font-size:13px;font-face:arial;width:100px">  </div>     
<BR>
<div style="display:inline-block; width:95px;text-align:left;">Şifre:</div> <div style="display:inline-block;"> <input type="password" size="15" name="password" maxlength="15" style="font-size:13px;font-face:arial;width:100px"> </div>
                           
<div style="margin-top:10px"><input type="submit"  name="submit" class="submit" value="GİRİŞ"></div>    

</div>  
</td></tr></table>      
             
<BR>
 MEGA E-Ticaret Sistemleri 


</div>    
</center>
</body>
</html>


