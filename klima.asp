<!--#include file='charset.asp'-->

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">  <html> 
<!--#include file='db.asp'-->

<head>
<meta name="viewport" content="width=device-width, initial-scale=1"> 
     
  <%   
    g= 19
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
seoDescription = "Antalya'da " & anas("isim") & " cozumleri icin Ozum Klima'dan kesif, montaj ve teknik servis destegi alin."

if t<>"" then
seoTitle = tipi("isim") & " " & anas("isim") & " | Ozum Klima Antalya"
seoDescription = "Antalya'da " & tipi("isim") & " " & anas("isim") & " cozumleri icin profesyonel kesif, uygulama ve servis destegi sunuyoruz."
end if

canonicalUrl = "https://ozumklima.com/klima.asp"
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

<div style="float:right">     


  <%  

set tipler = Server.CreateObJect("ADODB.RecordSet")
Sorgula = "Select * From tip where grup = "& g &"  order by sira desc"
tipler.open Sorgula,baglanti,1,3  
%>     
       

<%
Do while not tipler.Eof                                                                                                                                                                                                                  
%>   

<a href="?t=<%=tipler("id") %>" > <%=tipler("isim") %></a> &nbsp;&nbsp; |  &nbsp;&nbsp;
    

<%  
tipler.MoveNExt
Loop
%>

</div> 
 
<div style="clear:both;"></div>       




</Div>

<section class="aeoCevap" aria-labelledby="klima-aeo-baslik">
  <h1 id="klima-aeo-baslik">Antalya Klima Sat&#305;&#351;, Ke&#351;if ve Montaj Hizmeti</h1>
  <p><strong>K&#305;sa cevap:</strong> &Ouml;z&uuml;m Klima, Antalya'da ev, ofis ve ticari alanlar i&#231;in klima se&#231;imi, Mitsubishi Electric klima &#231;&ouml;z&uuml;mleri, ke&#351;if, montaj, bak&#305;m ve teknik servis deste&#287;i sunar.</p>
  <p>Antalya'da klima se&#231;erken yaln&#305;zca metrekareye bakmak yeterli de&#287;ildir. Cephe y&ouml;n&uuml;, g&uuml;ne&#351; alma s&uuml;resi, tavan y&uuml;ksekli&#287;i, cam y&uuml;zeyi, yal&#305;t&#305;m durumu, ki&#351;i say&#305;s&#305; ve cihaz&#305;n g&uuml;nl&uuml;k kullan&#305;m yo&#287;unlu&#287;u birlikte de&#287;erlendirilmelidir. Bu nedenle do&#287;ru klima tercihi genellikle yerinde ke&#351;if ile ba&#351;lar.</p>
  <h2>Klima se&#231;iminde en kritik konu nedir?</h2>
  <p>En kritik konu, mek&acirc;na uygun kapasite ve do&#287;ru montaj noktas&#305;d&#305;r. Gere&#287;inden k&uuml;&#231;&uuml;k cihaz konforu sa&#287;layamaz; gere&#287;inden b&uuml;y&uuml;k cihaz ise gereksiz enerji t&uuml;ketimi ve verimsiz &#231;al&#305;&#351;ma riski olu&#351;turabilir.</p>
  <h2>Hangi b&ouml;lgelerde klima hizmeti verilir?</h2>
  <p>&Ouml;z&uuml;m Klima, Antalya merkez ba&#351;ta olmak &uuml;zere D&ouml;&#351;emealt&#305;, Kepez, Muratpa&#351;a ve Konyaalt&#305; b&ouml;lgelerinde klima, ke&#351;if, montaj ve servis talepleri i&#231;in destek sa&#287;lar. G&uuml;ncel ke&#351;if ve servis uygunlu&#287;u i&#231;in <a href="iletisim.asp" title="&Ouml;z&uuml;m Klima ileti&#351;im">ileti&#351;im</a> sayfas&#305;ndan firmaya ula&#351;&#305;labilir.</p>
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
                                        
                                            
      <div class="yildiz">
<% if object("renk1") <> "FFFFFF" then %> <div style="background-color:#<%=object("renk1")%>">&nbsp;</div>  <% end if %>   
<% if object("renk2") <> "FFFFFF" then %><div style="background-color:#<%=object("renk2")%>">&nbsp;</div><% end if %>  
<% if object("renk3") <> "FFFFFF" then %><div style="background-color:#<%=object("renk3")%>">&nbsp;</div><% end if %>  
<% if object("renk4") <> "FFFFFF" then %><div style="background-color:#<%=object("renk4")%>">&nbsp;</div>  <% end if %>  
<% if object("renk5") <> "FFFFFF" then %><div style="background-color:#<%=object("renk5")%>">&nbsp;</div> <% end if %>  
</div>

     
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