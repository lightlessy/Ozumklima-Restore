

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

<meta name="author" content="Mega Tasarim"/>   

<link rel="icon" type="image/png" href="fav.png" />   


<link href="css/font.css" rel="stylesheet" type="text/css"/>   

<script src="js/jquery-2.1.1.js"></script>  
<script src="js/jquery-ui.min.js"></script>               

   <link href="css/fontawesome-all.css" rel="stylesheet">


           <link href="css/lity.css" rel="stylesheet">

<script src="js/lity.min.js"></script>

<style>
.hizmetLandingBlock{margin-top:24px;padding:18px;border:1px solid #E1E1E1;background:#FFFFFF;white-space:normal;box-sizing:border-box;font-family:Arial,Helvetica,sans-serif}.hizmetLandingBlock h2{margin:0 0 10px 0;font-size:1.35em;line-height:1.35em;color:#222}.hizmetLandingBlock p{margin:0 0 16px 0;line-height:1.65em;color:#444}.hizmetLandingGrid{display:flex;flex-wrap:wrap;gap:12px}.hizmetLandingCard{flex:1 1 220px;min-width:210px;border:1px solid #E7E7E7;background:#FAFAFA;padding:14px;box-sizing:border-box}.hizmetLandingCard b{display:block;margin-bottom:7px;color:#222;font-size:1.03em}.hizmetLandingCard span{display:block;line-height:1.55em;color:#555;margin-bottom:10px}.hizmetLandingCard a{color:#DE0303;font-weight:bold;text-decoration:none}.hizmetLandingCard a:hover{text-decoration:underline}.rehberMiniBlock{margin-top:18px;padding:14px;border:1px solid #E1E1E1;background:#F8F8F8;white-space:normal;line-height:1.65em}.rehberMiniBlock a{font-weight:bold;color:#3F3F3F;text-decoration:none}.rehberMiniBlock a:hover{text-decoration:underline}
</style>
   

  
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

   <div class="hizmetLandingBlock">
      <h2>Antalya Iklimlendirme Hizmetleri</h2>
      <p>Ozum Klima; Antalya'da klima servisi, montaj, bakim, Mitsubishi Electric klima cozumleri ve VRF sistemleri icin kesif odakli hizmet sayfalari sunar. Ihtiyaciniza en yakin basligi secerek detayli bilgi alabilirsiniz.</p>
      <div class="hizmetLandingGrid">
        <div class="hizmetLandingCard">
          <b>Antalya Klima Servisi</b>
          <span>Klima kesif, montaj, bakim ve teknik servis surecleri.</span>
          <a href="antalya-klima-servisi.asp">Detaylari incele</a>
        </div>
        <div class="hizmetLandingCard">
          <b>Dosemealti Klima Servisi</b>
          <span>Dosemealti villa, konut ve isletmeler icin klima destek sureci.</span>
          <a href="dosemealti-klima-servisi.asp">Detaylari incele</a>
        </div>
        <div class="hizmetLandingCard">
          <b>Antalya Klima Montaji</b>
          <span>Ic unite, dis unite, drenaj ve montaj noktasi planlamasi.</span>
          <a href="antalya-klima-montaji.asp">Detaylari incele</a>
        </div>
        <div class="hizmetLandingCard">
          <b>Antalya Klima Bakimi</b>
          <span>Sezon oncesi bakim, filtre, drenaj ve calisma kontrolleri.</span>
          <a href="antalya-klima-bakimi.asp">Detaylari incele</a>
        </div>
        <div class="hizmetLandingCard">
          <b>Mitsubishi Electric Klima</b>
          <span>Antalya'da Mitsubishi Electric klima secimi, kesif ve montaj.</span>
          <a href="antalya-mitsubishi-electric-klima.asp">Detaylari incele</a>
        </div>
        <div class="hizmetLandingCard">
          <b>Antalya VRF Sistemleri</b>
          <span>Otel, ofis, klinik ve ticari alanlar icin VRF cozumleri.</span>
          <a href="antalya-vrf-sistemleri.asp">Detaylari incele</a>
        </div>
      </div>
   </div>

   <div class="rehberMiniBlock">
      <b>Karar asamasinda misiniz?</b> Klima kapasitesi, kesif ve sistem secimi hakkinda bilgi almak icin <a href="klima-rehberi.asp">Klima Rehberi</a> sayfasini inceleyin.
   </div>

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