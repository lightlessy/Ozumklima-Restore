<!--#include file='charset.asp'-->

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">  <html> 
<!--#include file='db.asp'-->

<head>
<meta name="viewport" content="width=device-width, initial-scale=1"> 
     
  <%   
    g= 23
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

canonicalUrl = "https://ozumklima.com/havalandirma.asp"
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

<section class="aeoCevap" aria-labelledby="havalandirma-aeo-baslik">
  <h1 id="havalandirma-aeo-baslik">Antalya Havaland&#305;rma Sistemleri ve &#304;&#231; Hava Kalitesi &#199;&ouml;z&uuml;mleri</h1>
  <p> Havaland&#305;rma sistemleri, kapal&#305; alanlarda temiz hava giri&#351;i, kirli havan&#305;n uzakla&#351;t&#305;r&#305;lmas&#305; ve i&#231; hava kalitesinin korunmas&#305; i&#231;in kullan&#305;lan teknik &#231;&ouml;z&uuml;mlerdir. &Ouml;z&uuml;m Klima, Antalya'da konut, ofis, restoran, ma&#287;aza ve ticari alanlar i&#231;in havaland&#305;rma ihtiya&#231;lar&#305;na y&ouml;nelik destek sunar.</p>
  <p>Antalya gibi s&#305;cak ve nemli iklimlerde havaland&#305;rma yaln&#305;zca konfor konusu de&#287;ildir. Restoran, ofis, klinik, ma&#287;aza ve kapal&#305; ticari alanlarda hava sirk&uuml;lasyonu, koku kontrol&uuml;, nem dengesi ve kullan&#305;c&#305; konforu birlikte de&#287;erlendirilmelidir.</p>
  <h2>Havaland&#305;rma sistemi se&#231;erken nelere bak&#305;lmal&#305;?</h2>
  <p>Mek&acirc;n&#305;n kullan&#305;m amac&#305;, ki&#351;i yo&#287;unlu&#287;u, kapal&#305; alan hacmi, mevcut klima sistemi, taze hava ihtiyac&#305;, egzost gereksinimi ve kanal g&uuml;zerg&acirc;h&#305; birlikte incelenmelidir. Yanl&#305;&#351; tasarlanan sistem, yeterli hava de&#287;i&#351;imi sa&#287;layamayabilir.</p>
  <h2>&Ouml;z&uuml;m Klima hangi s&uuml;re&#231;lerde destek verir?</h2>
  <p>&Ouml;z&uuml;m Klima, Antalya'da havaland&#305;rma ke&#351;fi, sistem se&#231;imi, uygulama, montaj ve teknik servis konular&#305;nda destek sa&#287;lar. Havaland&#305;rma sistemi veya i&#231; hava kalitesi ihtiyac&#305; i&#231;in <a href="iletisim.asp" title="&Ouml;z&uuml;m Klima havaland&#305;rma ileti&#351;im">ileti&#351;im</a> sayfas&#305;ndan firmaya ula&#351;&#305;labilir.</p>
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