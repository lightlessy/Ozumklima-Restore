

<%  id=request.querystring("id") %>


<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html>
<head>  
 
<META HTTP-EQUIV="Content-Type" content="text/html; charset=iso-8859-9"> 
   <meta name="viewport" content="width=device-width, initial-scale=1">
 
 <!--#include file='db.asp'-->
<%
id = request.querystring("id")
if id = "" or isnumeric(id)=false then
   set varsayilanSayfa = Server.CreateObJect("ADODB.RecordSet")
   Sorgu = "Select top 1 * From sayfalar where ana=2 and yayin=1 order by sira asc"
   varsayilanSayfa.open Sorgu,baglanti,1,3
   if not varsayilanSayfa.eof then
      id = varsayilanSayfa("id")
   else
      id = 1
   end if
end if

set icerisi2 = Server.CreateObJect("ADODB.RecordSet")
Sorgu = "Select * From sayfalar where id="& id  &" and yayin=1"
icerisi2.open Sorgu,baglanti,1,3

if not icerisi2.eof then
   seoTitle = icerisi2("isim") & " Hizmetleri | Ozum Klima Antalya"
   seoDescription = "Ozum Klima'nin Antalya'da sundugu " & icerisi2("isim") & " hizmetleri ile kesif, uygulama, montaj ve teknik servis sureclerinde profesyonel destek alin."
else
   seoTitle = "Hizmetler | Ozum Klima Antalya"
   seoDescription = "Ozum Klima, Antalya'da iklimlendirme hizmetlerinde kesif, uygulama, montaj ve teknik servis destegi sunar."
end if
%>
     
<title><%=Server.HTMLEncode(seoTitle)%></title>

<meta name="description" content="<%=Server.HTMLEncode(seoDescription)%>" />
<link rel="canonical" href="https://ozumklima.com/h.asp?id=<%=id%>" />

<meta name="author" content="Mega Tasar�m"/>   

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
     <%=icerisi2("icerik") %>

   <div style="margin-top:22px; padding:14px; border:1px solid #E1E1E1; background:#F8F8F8; white-space:normal;">
      <b>Uygulama Ornekleri:</b> Antalya'da tamamladigimiz gercek saha uygulamalarini inceleyin.
      <br>
      <a href="uygulama-ornekleri.asp" style="display:inline-block; margin-top:8px; font-weight:bold; color:#3F3F3F;" onclick="analyticsTrack('case_study_cta_click', {page: 'hizmet', cta: 'uygulama_ornekleri_link'});">Uygulama orneklerini goruntule</a>
   </div>
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