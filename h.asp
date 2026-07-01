

<%  id=request.querystring("id") %>


<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html>
<head>  
 
<META HTTP-EQUIV="Content-Type" content="text/html; charset=iso-8859-9"> 
   <meta name="viewport" content="width=device-width, initial-scale=1">
 
 <!--#include file='db.asp'-->
     
    <%
    set icerisi2 = Server.CreateObJect("ADODB.RecordSet")
Sorgu = "Select * From sayfalar where id="& id  &" and yayin=1"
icerisi2.open Sorgu,baglanti,1,3   
%>      


<title><%=icerisi2("isim") %> Antalya   | Antalya Klima - Özüm Klima Mitsubishi Klima Antalya</title>    

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

   
   
  
<%
   id = request.querystring("id") 
if id="" then    

 set icerisi = Server.CreateObJect("ADODB.RecordSet")
Sorgu = "Select * From sayfalar where ana=2 and yayin=1 order by sira asc"
icerisi.open Sorgu,baglanti,1,3 

id= icerisi("id")  
else 
id = request.querystring("id") 
end if

%>    
                 

<Div class="tepeRestas"    style="background-image:url('images/arka-hizmet.jpg')"></Div>


 
<DIV class="MenuVeicerik">  

 <Div class="solMenu"> <!--  #include file='sol-hizmet.asp' --> </Div>
       


                         
  <Div  class="icerik2"  style="white-space: pre-line ; "><span><B>  <%=icerisi2("isim") %>  </B> </span>  
       <%=icerisi2("icerik") %>  <a href="http://www.ozumklima.com">Antalya klima </a> &nbsp;&nbsp;&nbsp; <a href="http://www.ozumklima.com/h.asp">Antalya klima fiyatları</a>  &nbsp;&nbsp;&nbsp; <a href="http://www.ozumklima.com/h.asp?h=mitsubishi-elektrik-klima-yetkili-bayii&id=12">Mitsubishi Antalya klima </a>  &nbsp;&nbsp;&nbsp;  <a href="http://www.ozumklima.com">Antalya Mitsubishi Klima Bayii Servis </a>   &nbsp;&nbsp;&nbsp; <a href="http://www.ozumklima.com/h.asp?h=mitsubishi-elektrik-klima-yetkili-servisi&id=13">Antalya klima servisi </a>   &nbsp;&nbsp;&nbsp;  <a href="http://www.ozumklima.com/vrf_sistemler.asp">Antalya havalandırma  </a>
  </Div>
       
       
                 
  
  
  
              
 <div style="clear:both"  ></div>
</DIV>             
         
     
     
     
     
                  
     
            <script>
$(window).load(function(){       


var anchor=$(location).attr('hash');  
var modeller = $( ".icerik2" );
var offset = modeller.offset();
    



   
if ( anchor =="#hak" ) {   $("html, body").animate({ scrollTop: offset.top-150 }, 600);           }    
else {   } 




})       

       </script>



       
         

             
            
                 

    <!--  #include file='alt.asp' -->      

   

</body>
</html>