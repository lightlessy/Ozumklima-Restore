<!--#include file='charset.asp'-->

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html>
<head>  
 
<meta http-equiv="Content-Type" content="text/html; charset=windows-1254"> 
<meta name="viewport" content="width=device-width, initial-scale=1">

<title>&Ouml;z&uuml;m Klima &#304;leti&#351;im | Antalya Klima Servisi ve Ke&#351;if Talebi</title>   

<meta name="description" content="&Ouml;z&uuml;m Klima ile ileti&#351;ime ge&#231;in. Antalya'da klima, VRF, havaland&#305;rma, ke&#351;if, teklif, montaj, bak&#305;m ve servis deste&#287;i al&#305;n.">
<link rel="canonical" href="https://ozumklima.com/iletisim.asp">

<meta name="author" content="Mega Tasar&#305;m">   

<link rel="icon" type="image/png" href="fav.png">   

<link href="css/font.css" rel="stylesheet" type="text/css">   

<script src="js/jquery-2.1.1.js"></script>  
<script src="js/jquery-ui.min.js"></script>               

<link href="css/fontawesome-all.css" rel="stylesheet">

<% hform= request.querystring("hform") %>
</head>

<body> 

<!--#include file='ust.asp'-->

<DIV>
               
<Div class="slogan" style="height:20px"></Div>

<Div class="iletTas">  

   <Div style="float:left; max-width:300px" class="adresTel"> 
   
      <font style="font-size:1.3em; color:#FEAD41;">Alo Konfor Hatt&#305;<BR></font>
      <span class="alttel"><i class="fas fa-phone"></i>  <%=ayars("tel") %>  <BR><BR></span>
                  
      <B>Adres: </B><BR>
      <%=ayars("adres") %>

      <BR><BR>
                   
      <B>E-Posta:</B><BR>
      <a href="mailto:<%=ayars("eposta") %>" target=_Blank style="color:#000000"> <%=ayars("eposta") %></a> <BR><BR>
                  
      <B>Faks: </B><BR>
      <%=ayars("faks") %> <BR><BR> 
                  
      <B>Bizi Takip Edin: </B><BR>
      <a href="<%=ayars("face")%>" target=_blank><i class="fab fa-facebook-f"></i></a>&nbsp;&nbsp;  
      <a href="<%=ayars("insta")%>" target=_blank><i class="fab fa-instagram"></i></a>&nbsp;&nbsp;
           
   </Div>

   <Div style="float:right;" class="formTas">  
     
      <div style="text-align:center;">

         <form method="post" name="cic" action="iletisim-send.asp" class="form"
      onsubmit="safeTrack('contact_form_submit_attempt', {source:'iletisim_page'});">          
                          
            <input name="isim" type="text" placeholder="&#304;sim" size="26" maxlength="60" required>
            <input name="mail" type="email" placeholder="E-Posta Adresiniz" size="26" maxlength="60" required>  
            <input name="tel" type="text" placeholder="Telefon Numaran&#305;z" size="10" maxlength="11" onkeypress="return numbersonly(this, event)"> 
            <textarea name="mes" cols="25" rows="6" placeholder="Mesaj&#305;n&#305;z" required></textarea>	 
            <input type="submit" name="submit" value="G&ouml;nder" class=btn style="color:#FFFFFF"> 	 	 
      
         </form>         
      
      </div>   
         
   </Div>

   <Div style="clear:both"></Div>   

</Div>            

</DIV>        

<%=ayars("map") %>                                                                                         
      
<!--  #include file='alt.asp' -->      

</body>
</html>