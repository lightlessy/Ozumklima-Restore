<!--#include file='charset.asp'-->

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">  <html> 
<!--#include file='db.asp'-->

<head>
<meta name="viewport" content="width=device-width, initial-scale=1"> 
     
  <%   
    g= 20
 t= request.querystring("t")
 s= request.querystring("s")       
   d= request.querystring("d") 
   
   if d="" then
   diz="sira asc"  
   elseif d="0" then    
   diz="sira desc"    
    elseif d="19" then    
   diz="fiyat asc"  
   elseif d="91" then    
   diz="fiyat desc"
   end if       
   
   
 
    set anas = Server.CreateObJect("ADODB.RecordSet")
Sorgula = "Select * From anagrup where id= " & g &" "
anas.open Sorgula,baglanti,1,3        

if t<>"" then 
set tipi = Server.CreateObJect("ADODB.RecordSet")
Sorgula = "Select * From tip where id= " & t &" "
tipi.open Sorgula,baglanti,1,3    
end if

if s<>"" then   
set sorf = Server.CreateObJect("ADODB.RecordSet")
Sorgula = "Select * From surface where id= " & s &" "
sorf.open Sorgula,baglanti,1,3   
end if 

seoTitle = anas("isim") & " Cozumleri | Ozum Klima Antalya"
seoDescription = "Antalya'da " & anas("isim") & " sistemleri icin Ozum Klima'dan kesif, montaj ve teknik servis destegi alin."

if t<>"" then
seoTitle = tipi("isim") & " " & anas("isim") & " | Ozum Klima Antalya"
seoDescription = "Antalya'da " & tipi("isim") & " " & anas("isim") & " cozumleri icin profesyonel kesif, uygulama ve servis destegi sunuyoruz."
end if

canonicalUrl = "https://ozumklima.com/vrf_sistemler.asp"
if t<>"" then
canonicalUrl = canonicalUrl & "?t=" & Server.URLEncode(t)
if s<>"" then
canonicalUrl = canonicalUrl & "&s=" & Server.URLEncode(s)
end if
elseif s<>"" then
canonicalUrl = canonicalUrl & "?s=" & Server.URLEncode(s)
end if
%>


<title><%=Server.HTMLEncode(seoTitle)%></title>

  <meta name="description" content="<%=Server.HTMLEncode(seoDescription)%>" />
  <link rel="canonical" href="<%=canonicalUrl%>" />


 <meta http-equiv="Content-Type" content="text/html; charset=windows-1254"> 
  
    <link href='https://fonts.googleapis.com/css?family=Source+Sans+Pro:300' rel='stylesheet' type='text/css'>   
       
      <link rel="stylesheet" href="css/font.css">

      <link rel="shortcut icon" href="images/favicon.png" >      

   <link href="css/fontawesome-all.css" rel="stylesheet">
      <script src="js/jquery-2.1.1.js"></script>
      
<style>
.aeoCevap {width:92%; max-width:1120px; margin:24px auto 30px auto; padding:26px 4%; background:#FFFFFF; border:1px solid #E8E8E8; box-sizing:border-box; font-family:Arial, Helvetica, sans-serif; color:#2B2B2B;}
.aeoCevap h1 {margin:0 0 12px 0; padding:0; font-size:1.55em; line-height:1.35em; color:#171717;}
.aeoCevap h2 {margin:18px 0 7px 0; padding:0; font-size:1.12em; line-height:1.4em; color:#222222;}
.aeoCevap p {margin:0 0 12px 0; padding:0; font-size:1em; line-height:1.68em; color:#3A3A3A;}
.aeoCevap ul {margin:8px 0 0 18px; padding:0; line-height:1.65em;}
.aeoCevap a {color:#DE0303; font-weight:bold; text-decoration:none;}
.aeoCevap a:hover {text-decoration:underline;}
</style>

</head>

<body>
       
             
       
<!--#include file='ust.asp'-->     


 <Div  class="navi" >

<div style="float:left">
<a  href="?g=<%=g%>" > <%=anas("isim") %> </a> <% if t<>"" then %>   /  <a  href="?g=<%=g%>&t=<%=t%>" ><%=tipi("isim") %> </a> <% end if %>  <% if s<>"" then %>  /  <a  href="#" ><%=sorf("isim") %> </a> <% end if %>                                                                                                                                                                                                                                                      
</div>



<div style="clear:both;"></div>       


   
</Div>

<section class="aeoCevap" aria-labelledby="vrf-aeo-baslik">
  <h1 id="vrf-aeo-baslik">Antalya VRF Sistemleri ve Merkezi &#304;klimlendirme &#199;&ouml;z&uuml;mleri</h1>
  <p><strong>K&#305;sa cevap:</strong> VRF sistemleri; otel, ofis, klinik, restoran, ma&#287;aza ve &#231;ok odal&#305; ticari yap&#305;larda farkl&#305; alanlar&#305;n tek bir merkezi sistemle verimli &#351;ekilde iklimlendirilmesini sa&#287;layan profesyonel &#231;&ouml;z&uuml;mlerdir.</p>
  <p>Antalya'da VRF sistemi se&#231;imi, yaln&#305;zca cihaz kapasitesine bakarak yap&#305;lmamal&#305;d&#305;r. Kat plan&#305;, i&#231; &uuml;nite yerle&#351;imi, d&#305;&#351; &uuml;nite konumu, borulama mesafesi, drenaj, kullan&#305;m senaryosu ve enerji verimlili&#287;i birlikte projelendirilmelidir. Bu nedenle VRF uygulamalar&#305;nda ke&#351;if ve teknik proje a&#351;amas&#305; kritik rol oynar.</p>
  <h2>VRF sistemi kimler i&#231;in uygundur?</h2>
  <p>VRF sistemleri, birden fazla b&ouml;l&uuml;m&uuml;n ayr&#305; ayr&#305; kontrol edilmesi gereken yap&#305;lar i&#231;in uygundur. Ofis katlar&#305;, oteller, klinikler, showroomlar, restoranlar ve ma&#287;azalar bu sisteme en s&#305;k ihtiya&#231; duyan alanlard&#305;r.</p>
  <h2>&Ouml;z&uuml;m Klima VRF s&uuml;recinde ne yapar?</h2>
  <p>&Ouml;z&uuml;m Klima, Antalya'da VRF ihtiya&#231;lar&#305; i&#231;in ke&#351;if, kapasite planlama, cihaz se&#231;imi, uygulama, montaj ve teknik servis s&uuml;re&#231;lerinde destek sunar. VRF ke&#351;if veya teklif talebi i&#231;in <a href="iletisim.asp" title="&Ouml;z&uuml;m Klima VRF ileti&#351;im">ileti&#351;im</a> sayfas&#305;ndan firmaya ula&#351;&#305;labilir.</p>
</section>
 
     
   <script>
     $(document).ready(function(){

      $(".goturen").on('change', function () {
          var url = $(this).val(); // get selected value
          if (url) { // require a URL
              window.location = url; // redirect
          }
          return false;
      });     
      
      
    });
</script>
 
    
     <Div style="width:100%; max-width:1200px; margin:0px auto; ">

    
 <BR><BR>



 <DIV   class="kutularic" >         
      
      
<%  

if t<>"" then
bul1="and tip="&t&""    
end if
           
if s<>"" then
bul2="and surface="&s&""    
end if

set object = Server.CreateObJect("ADODB.RecordSet")
Sorgula = "Select * From products where anagrup ="& g &" "& bul1 &"  "& bul2 &" and yayin=1 order by "&diz&"  "
object.open Sorgula,baglanti,1,3       


Do while not object.Eof     
%>      
      
      
      
      
      
                     <Div>   
                       
           
                                     
                                 <a href="<%=cevir(object("isim"))%>-<%=object("AffiliateID")%>">  
 <div class="urRes"> <img src="urunler/<%=object("foto1")%>" width=100%>
                                     <% if object("yeni")=1 then %><div class="yeni"><img  src="images/yeni.png"></div> <% end if %>
                                   <% if object("ozel5")=1 then %><div class="yeni"><img  src="images/indi.png"></div>   <% end if %>
                                     </div>  
                                                      
                                     
         <div  class="urAd" ><%=object("isim")%></div>   
                                        
                                            


     
<div class="detayal">&#304;ncele </div>    
            

 
  </a>   

</Div>
                       
                       
           
           <%  
object.MoveNExt
Loop
%>                    
                       
                       

      </DIV>     
    
      
      </div>
     
<div style="clear:both"></div>
      
        
           </Div>  
           
           
           
           <BR><BR><BR>&nbsp;
 
<!--#include file='alt.asp'-->
 

    
</body>

</html>