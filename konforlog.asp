

<%  id=request.querystring("id") %>


<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html>
<head>  
 
<META HTTP-EQUIV="Content-Type" content="text/html; charset=iso-8859-9"> 
   <meta name="viewport" content="width=device-width, initial-scale=1">
  

<title>Konforlog | Antalya Klima - Özüm Klima Mitsubishi Klima Antalya</title>    

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
Sorgu = "Select * From sayfalar where ana=3 and yayin=1 order by sira asc"
icerisi.open Sorgu,baglanti,1,3 

id= icerisi("id")  
else 
id = request.querystring("id") 
end if


 set icerisi = Server.CreateObJect("ADODB.RecordSet")
Sorgu = "Select * From sayfalar where id="& id  &" and yayin=1"
icerisi.open Sorgu,baglanti,1,3   
%>                       

<Div class="tepeRestas"    style="background-image:url('images/arka-kurumsal.jpg')"></Div>


 
<DIV class="MenuVeicerik">  

 <Div class="solMenu"> <!--  #include file='sol-blog.asp' --> </Div>
      

  <Div  class="icerik2"  style="white-space: pre-line ;"><span><B>  <%=icerisi("isim") %>  </B> </span>  
       <%=icerisi("icerik") %>
  </Div>     
       



       
       
                 
  
  
  
              
 <div style="clear:both"  ></div>
</DIV>             
         
     


    <!--  #include file='alt.asp' -->      

   

</body>
</html>