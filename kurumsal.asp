

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

 <%   
 id = request.querystring("id") 
if id="" then    

 set icerisi = Server.CreateObJect("ADODB.RecordSet")
Sorgu = "Select * From sayfalar where ana=1 and yayin=1 order by sira asc"
icerisi.open Sorgu,baglanti,1,3 

id= icerisi("id")  
else 
id = request.querystring("id") 
     
if id=11 then
response.redirect("referanslar.asp")
end if 

end if


 set icerisi = Server.CreateObJect("ADODB.RecordSet")
Sorgu = "Select * From sayfalar where id="& id  &" and yayin=1"
icerisi.open Sorgu,baglanti,1,3   
%>                       

<Div class="tepeRestas"    style="background-image:url('images/arka-kurumsal.jpg')"></Div>


 
<DIV class="MenuVeicerik">  

 <Div class="solMenu"> <!--  #include file='sol-kurumsal.asp' --> </Div>
      

  <Div  class="icerik2"  style="white-space: pre-line ;"><span><B>  <%=icerisi("isim") %>  </B> </span>  
       <%=icerisi("icerik") %>  
       
       
  
  
  <% if id=20  then %>
  
<B>  Yorumlarını Bekliyoruz: </B>
  <FORM METHOD="POST" action="yorumla.asp"  class="form">         
  <TEXTAREA   style="HEIGHT: 90px"  name="yorum" placeholder="Yorumunuzu Buraya Yazın..." required></TEXTAREA> 
 <INPUT  type="text"  name="isim" placeholder="Ad Soyad" required>   
<input type="submit" value="Gönder" class=btn  style="color:#FFFFFF"> 
 </FORM>          
<hr style="height: 1px;background-color: #B9B9B9; border:none;">        
  <%
set yorumx = Server.CreateObJect("ADODB.RecordSet")
Sorgu = "Select * From yorumlar where onay=1 order by id desc"
yorumx.open Sorgu,baglanti,1,3   

Do while not yorumx.Eof                                                                                                                                                                                                                  
%> 
        
<%=yorumx("yorum") %> <BR>
<I><%=yorumx("isim") %> </I>
<span style="font-size:1em; line-height:2em; color:#BF0060"><%=yorumx("cevap")%> </span>  
<hr style="height: 1px;background-color: #B9B9B9; border:none;">      
<%  
yorumx.MoveNExt
Loop
%>
             

  <% end if %>
      
       
       
  </Div>     
       



       
       
                 
  
  
  
              
 <div style="clear:both"  ></div>
</DIV>             
         
     


    <!--  #include file='alt.asp' -->      

   

</body>
</html>