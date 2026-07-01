<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html>
<head>  
 
<META HTTP-EQUIV="Content-Type" content="text/html; charset=iso-8859-9"> 
   <meta name="viewport" content="width=device-width, initial-scale=1">


<title>İletişim  | Antalya Klima - Özüm Klima Mitsubishi Klima Antalya</title>   

<meta name="author" content="Mega Tasarım"/>   

<link rel="icon" type="image/png" href="fav.png" />   

<link href="css/font.css" rel="stylesheet" type="text/css"/>   

<script src="js/jquery-2.1.1.js"></script>  
<script src="js/jquery-ui.min.js"></script>               

   <link href="css/fontawesome-all.css" rel="stylesheet">


  <% hform= request.querystring("hform") %>
</head>

<body> 

  
<!--  #include file='ust.asp' -->


<DIV >
               

<Div class="slogan" style="height:20px"></Div>

<Div  class="iletTas">  


   <Div  style="float:left; max-width:300px"    class="adresTel"  > 
   
                   <font style="font-size:1.3em; color:#FEAD41 ;">Alo Konfor Hattı<BR></font>
                  <span class="alttel"><i class="fas fa-phone"></i>  <%=ayars("tel") %>  <BR><BR></span>
                  
                   <B>Adres: </B><BR>
<%=ayars("adres") %>

<BR><BR>
                   
                  <B>E-Posta:</B>      <BR>
         <a href="mailto:<%=ayars("eposta") %>" target=_Blank style="color:#000000"> <%=ayars("eposta") %></a> <BR>   <BR>
                  
                  <B>Faks: </B><BR>
       <%=ayars("faks") %>   <BR><BR> 
                  
                   <B>Bizi Takip Edin: </B><BR>
                  <a href="<%=ayars("face")%>" target=_blank><i class="fab fa-facebook-f"></i> </a>&nbsp;&nbsp;  
                       <a href="<%=ayars("insta")%>" target=_blank><i class="fab fa-instagram"></i> </a>&nbsp;&nbsp;
           

                  
            </Div>

         
     <Div  style="float:right;"   class="formTas"   >  
     
<div style="text-align:center;  ">

               
 <form method="post" name="cic"  action="iletisim-send.asp"      class="form">          
                          

<input name="isim" type="text"  placeholder="İsim"  size="26" maxlength="60" required >
<input name="mail" type="email"      placeholder="E-Posta Adresiniz"  size="26" maxlength="60"  required>  
<input name="tel" type="text"   placeholder="Telefon Numaranız "  size="10"   maxlength="11" onkeypress="return numbersonly(this, event)"> 
<textarea name="mes" cols="25" rows="6"  placeholder=" Mesajınız"  required  ></textarea>	 
  <input type="submit"  name="submit" value="Gönder"   class=btn  style="color:#FFFFFF"> 	 	 

      </form>         
      
      

               
         </div>   
         
  </Div>


<Div  style="clear:both" ></Div>   





</Div>            





 </DIV>        




 <%=ayars("map") %>                                                                                         
      
      
      
    <!--  #include file='alt.asp' -->      

   

</body>
</html>