

<%  id=request.querystring("id") %>


<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html>
<head>  
 
<META HTTP-EQUIV="Content-Type" content="text/html; charset=iso-8859-9"> 
   <meta name="viewport" content="width=device-width, initial-scale=1">
  

<title>Kurumsal | Antalya Klima - Özüm Klima Mitsubishi Klima Antalya</title>    

<meta name="author" content="Mega Tasarım"/>   

<link rel="icon" type="image/png" href="fav.png" />   



<link href="css/font.css" rel="stylesheet" type="text/css"/>   

<script src="js/jquery-2.1.1.js"></script>  
<script src="js/jquery-ui.min.js"></script>               
                                                                       




   <link href="css/fontawesome-all.css" rel="stylesheet">
   <link href="css/lity.css" rel="stylesheet">

<script src="js/lity.min.js"></script>


    
</head>

<body> 


<!--  #include file='ust.asp' -->

                

<Div class="tepeRestas"    style="background-image:url('images/arka-kurumsal.jpg')"></Div>


 
<DIV class="MenuVeicerik">  

 <Div class="solMenu"> <!--  #include file='sol-kurumsal.asp' --> </Div>
      

  <%
    set icerisi = Server.CreateObJect("ADODB.RecordSet")
Sorgu = "Select * From sayfalar where id=11 and yayin=1"
icerisi.open Sorgu,baglanti,1,3   
%>          

  <Div  class="icerik2"  ><span><B>  <%=icerisi("isim") %>   </B> </span> <%=icerisi("icerik") %>      </Div> 
<Div  class="tedaviler">    

 


  <%
set icerisi = Server.CreateObJect("ADODB.RecordSet")
Sorgu = "Select * From refers"
icerisi.open Sorgu,baglanti,1,3   

Do while not icerisi.Eof                                                                                                                                                                                                                  
%>       

   <a href="http://<%=icerisi("link") %> "   target=_blank     > 
          <Div style="text-align:center; padding:25px; color:#3F3F3F; line-height:2em">  
<img src="ref/<%=icerisi("resim") %>"  width=250>  <BR>
   <%=icerisi("isim") %>              
          </Div>
          </a>    


<%  
icerisi.MoveNExt
Loop
%>
                             


</Div>    


                 
  
  
  
              
 <div style="clear:both"  ></div>
</DIV>             
         
     


    <!--  #include file='alt.asp' -->      

   

</body>
</html>