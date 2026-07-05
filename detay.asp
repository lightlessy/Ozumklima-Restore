<!--#include file='charset.asp'-->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">  <html>

<head>     

  <!--#include file='db.asp'-->      
  
<meta name="viewport" content="width=device-width, initial-scale=1">

 <meta http-equiv="Content-Type" content="text/html; charset=windows-1254"> 
  
    <link href='https://fonts.googleapis.com/css?family=Source+Sans+Pro:300' rel='stylesheet' type='text/css'>   
    
      <link rel="stylesheet" href="css/font.css"><link rel="shortcut icon" href="fav.png" >      

      
   <link href="css/fontawesome-all.css" rel="stylesheet">
          
             <script src="js/jquery-2.1.1.js"></script>          

      
        <link rel="stylesheet" href="css/magnific-popup.css"> 
<script src="js/jquery.magnific-popup.js"></script>  
       
         
               <script src="js/starrr.js"></script>
        <link rel="stylesheet" href="css/starrr.css">
              
                  <link rel="stylesheet" href="css/lity.css">   
                  <script src="js/lity.min.js"></script>

<link rel="stylesheet" href="css/darkbox.css"/>
           
           
           <%   
  txt=Request.ServerVariables("QUERY_STRING")

sonYeri=InStrRev(txt,"/") 'son / isaretinin yeri  

sonuc=Mid(txt,sonYeri+1,Len(txt) )  'son / isaretinden kalan yazinin tamami   



istekbolum = Split(sonuc,"-")  'tirelerden bolerek dizi olustu  

uzunluk=uBound(istekbolum)     'tirelerden bolerek olusturdugumuz dizinin uzunlugu    
  
pidi=istekbolum(uzunluk)  
ad= request.querystring("ad")            
  

if ad= "" then
ad=1
end if

if isnumeric(ad)=false then
ad=1
end if
 
 
set uruns = Server.CreateObJect("ADODB.RecordSet")
Sorgula = "Select * From products where AffiliateID= "& pidi
uruns.open Sorgula,baglanti,1,3      
          
set anas = Server.CreateObJect("ADODB.RecordSet")
Sorgula = "Select * From anagrup where id= " & uruns("anagrup") &" "
anas.open Sorgula,baglanti,1,3        

set tipi = Server.CreateObJect("ADODB.RecordSet")
Sorgula = "Select * From tip where id= " & uruns("tip") &" "
tipi.open Sorgula,baglanti,1,3     

set sorf = Server.CreateObJect("ADODB.RecordSet")
Sorgula = "Select * From surface where id= " & uruns("surface") &" "
sorf.open Sorgula,baglanti,1,3  


Set Yonlendir = Server.CreateObject("Adodb.Recordset")
Sorgula = "Select * From products where AffiliateID= "& pidi
Yonlendir.Open Sorgula,baglanti,1,3
Yonlendir("hit") = Yonlendir("hit")+1
Yonlendir.Update  

seoTitle = uruns("isim") & " | " & anas("isim") & " Cozumu | Ozum Klima"

seoDescription = uruns("descr") & ""
seoDescription = Replace(seoDescription, vbCrLf, " ")
seoDescription = Replace(seoDescription, vbCr, " ")
seoDescription = Replace(seoDescription, vbLf, " ")
seoDescription = Replace(seoDescription, "&nbsp;", " ")

do while InStr(seoDescription, "<") > 0 and InStr(seoDescription, ">") > InStr(seoDescription, "<")
  basla = InStr(seoDescription, "<")
  bitir = InStr(basla, seoDescription, ">")
  seoDescription = Left(seoDescription, basla - 1) & " " & Mid(seoDescription, bitir + 1)
loop

seoDescription = Trim(seoDescription)
do while InStr(seoDescription, "  ") > 0
  seoDescription = Replace(seoDescription, "  ", " ")
loop

if seoDescription = "" then
  seoDescription = uruns("isim") & " icin Ozum Klima'dan Antalya'da kesif, montaj ve teknik servis destegi alin."
end if

if Len(seoDescription) > 160 then
  seoDescription = Left(seoDescription, 157) & "..."
end if

canonicalUrl = "https://ozumklima.com/" & cevir(uruns("isim")) & "-" & uruns("AffiliateID")
 


%>   


<title><%=Server.HTMLEncode(seoTitle)%></title>
                                                           
 
  <meta name="description" content="<%=Server.HTMLEncode(seoDescription)%>" />
  <link rel="canonical" href="<%=canonicalUrl%>" />


</head>

<body>
       
             
       
<!--#include file='ust.asp'-->
    
              
                    <Div  class="navSite" style="font-size:1.7em; " >
        <%=anas("isim") %>   /  <%=tipi("isim") %>
        </Div>
               
              
    <BR>
    
    <DIV  class="UrDetANA"  >
                
                  
                  <div  class="galeri" >     
                             
                             <div style="width:100%; height:350px;background-image:url('urunler/<%=uruns("foto1")%>')" class="devres"   data-darkbox="urunler/<%=uruns("foto1")%>"    >   </div>
                             <div class="thumb "  <% if uruns("foto2") ="" then %>  style="display:none"  <% end if %>>          
                             <% if uruns("foto2") <>"" then %> 
                    <img             src="urunler/<%=uruns("foto1")%>"              data-darkbox="urunler/<%=uruns("foto1")%>"                data-darkbox-description=""    />          <% end if %>
                <% if uruns("foto2") <>"" then %>                 <img             src="urunler/<%=uruns("foto2")%>"              data-darkbox="urunler/<%=uruns("foto2")%>"             data-darkbox-description=""  />    <% end if %>
                <% if uruns("foto3") <>"" then %>                 <img             src="urunler/<%=uruns("foto3")%>"              data-darkbox="urunler/<%=uruns("foto3")%>"             data-darkbox-description=""  />    <% end if %>
               <% if uruns("foto4") <>"" then %>                  <img             src="urunler/<%=uruns("foto4")%>"              data-darkbox="urunler/<%=uruns("foto4")%>"             data-darkbox-description=""  />     <% end if %>
               <% if uruns("foto5") <>"" then %>                  <img             src="urunler/<%=uruns("foto5")%>"              data-darkbox="urunler/<%=uruns("foto5")%>"             data-darkbox-description=""  />   <% end if %>
         </div>
                           
                  </div>
                               
                     
            <script src="js/darkbox.js"></script>
                  
                     

   
   <div class="fiyatsepet"> 
    
    <div  class="urunOz">
  <div  class="urunbaslik" ><%=uruns("isim")%> 
 
 
 </div>

<BR>

<span class="urunicerik"> <%=uruns("descr")%>    </span>
 
 <BR> <BR>
    
  <form method="POST" action="iletisim.asp" name="myForm"  >   

<input type="hidden" name="fiyat"   value="<% if uruns("indirim")=0 then %><%=uruns("fiyat") %><% else %><%=uruns("indirim") %><% end if %>" >           
<input type="hidden" name="uid"    value="<%=Session("CustomerID")%>">      
 <input type="hidden" name="product"    value="<%=uruns("AffiliateID")%>">      

<% if   uruns.fields("uretici")&"" <> "" then  %>   
<a href="http://www.ozumklima.com/urunler/<%=uruns("uretici") %>" target=_Blank><div class="whats" style="color:#3A3A3A"><img src="images/ikon8.png" >  <div><b>Teknik &Ouml;zellikler</B></div> </div></a>  
<% end if %>
                          
<% if   uruns.fields("dokuman")&"" <> "" then  %>   
<a href="http://www.ozumklima.com/urunler/<%=uruns("dokuman") %>" target=_Blank><div class="whats" style="color:#3A3A3A"><img src="images/ikon5.png" >  <div><b>Bro&#351;&uuml;r &#304;ndir</B></div> </div></a> 
<% end if %>

<BR><BR><BR>  
  <input  type="submit" value="&Uuml;cretsiz Ke&#351;if" name="update" id="sbmt" class="sepete" style="background:#A3C225;color:#FFFFFF; font-family:Source Sans Pro; font-weight:bold; font-size:1.2em; padding: 12px 32px 12px 32px; -webkit-border-radius: 6px; -moz-border-radius: 6px; border-radius: 6px;  " onclick="analyticsTrack('product_interest_click', {cta: 'ucretsiz_kesif', product_id: '<%=uruns("AffiliateID")%>'});">    
</form>

</div>    
</div>   

              
                  
<BR><BR><BR>      


              

<div  class="urunDetay"   >                                                     

                      
 <div  class="urunicerik">    
 

  <%=uruns("video")%>  
 <% if uruns("video") <>"" then %> <BR> <BR> <% end if %> 


 
 <%=uruns("notlar")%>                               
                         

<BR>   <BR>
<div style="text-align:center">

  <form method="POST" action="iletisim.asp" name="myForm"  >   
            <input  type="submit" value="BU &Uuml;R&Uuml;NLE &#304;LG&#304;LEN&#304;YORUM" name="update" id="sbmt" class="sepete" style="background:#A3C225;color:#FFFFFF; font-family:Source Sans Pro; font-weight:bold; font-size:1.2em; padding: 12px 32px 12px 32px; -webkit-border-radius: 6px; -moz-border-radius: 6px; border-radius: 6px;  " onclick="analyticsTrack('product_interest_click', {cta: 'bu_urunle_ilgileniyorum', product_id: '<%=uruns("AffiliateID")%>'});">    
</form>   
<BR><BR>&nbsp;                          
</div>


   </div> </div>
                  
        

  
  

	        
                                         
                                 
    </DIV>
                

            
 
 
 
        </DIV>
 
 



 
 
 
 
<% 
 
       set ilg = Server.CreateObJect("ADODB.RecordSet")
Sorgula = "Select * From related where urun = "& uruns("AffiliateID") &" "
ilg.open Sorgula,baglanti,1,3   
if not ilg.eof then 

i = 1 %>
 
     
                <!----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------->

<link href="css/owl.carousel.css" rel="stylesheet">
<link href="css/owl.theme.css" rel="stylesheet">  
                          
 
 
    <DIV   class="Box" >

     <span class="baslik"  style="border:none; ">Benzer &Uuml;r&uuml;nler</span>
      
  <BR><BR>
 <div id="owl-demo" class="owl-carousel" >
                  
    
          
                 <%


Do while not ilg.Eof  


set rurun = Server.CreateObJect("ADODB.RecordSet")
Sorgula = "Select * From products where AffiliateID=  "& ilg("ilgili") &" "
rurun.open Sorgula,baglanti,1,3        
   %>  
              
                      
                    
                       <div class="item" >
                                   
                                <a     href="<%=cevir(rurun("isim"))%>-<%=rurun("AffiliateID")%>">  
                                  
                                    <div class="urRes"  style="background-image:url('urunler/<%=rurun("foto1")%>');">
                                     <% if rurun("yeni")=1 then %><div class="yeni"><img  src="images/yeni.png"></div> <% end if %>
                                   <% if rurun("indirim")<>0 then %><div class="yeni"><img  src="images/indi.png"></div>   <% end if %>
                                     </div>        
                                       <div  class="urAd" style="color:#000000"><%=rurun("isim")%></div>
                                    <div  class="fiyat"  >
                                             <% if rurun("indirim")<>0 then %> <div class="yuzde">%<%= Round((rurun("fiyat") / rurun("indirim")*100)-100) %></div> <% end if %>
                                            <div style="text-align:left !important ;">
                                                 <% if rurun("indirim")<>0 then %>
                                                 <span class="indirimsiz"><%=FormatNumber(rurun("fiyat"),2)%> <i class="fas fa-lira-sign" aria-hidden="true"  style="font-size:.9em" ></i> </span>    
                                                 <%=FormatNumber(rurun("indirim"),2)%> <i class="fas fa-lira-sign" aria-hidden="true"  style="font-size:.7em" ></i>
                                                 <% else %>
                                               <%=FormatNumber(rurun("fiyat"),2)%> <i class="fas fa-lira-sign" aria-hidden="true"  style="font-size:.7em" ></i>   
                                               <% end if %>
                                            </div>  
                                            
 <BR>
<%  
set yorumcu = Server.CreateObJect("ADODB.RecordSet")
Sorgula = "Select * From yorumlar where urun ="& rurun("AffiliateID")&" and rating > 3 and onay =1"
yorumcu.open Sorgula,baglanti,1,3 
if not yorumcu.eof then   
%>
<% for i=1 to  yorumcu("rating") %>
<i class="fa fa-star" aria-hidden="true"  style="font-size:.7em; color:#F39F4B" ></i>    
<% next
 end if %>
 
 
                                    </div>
                                        
                                        </a>   
                                        
                                        
                </div>  
             
                                  
        <%
ilg.MoveNext
Loop  
%>
           
               

    </div>         
 
     </DIV>
    

     <script src="js/owl.carousel.min.js"></script>
 
 
    <!-- Demo -->
 
    <style>
 #owl-demo .item{   position:relative ; background:#ffffff;  display:inline-block;vertical-align:top ;width:43%;max-width:250px;   margin:0px 5px 25px 5px; border:1px solid #E4E4E4;  border-bottom:1px solid #E4E4E4 ; padding-bottom:23px;    }  
    
 
  #owl-demo .item .urRes:hover { opacity:.65 }
   
    .customNavigation{
      text-align: center;
    }   
    
  
    .customNavigation a{
      -webkit-user-select: none;
      -khtml-user-select: none;
      -moz-user-select: none;
      -ms-user-select: none;
      user-select: none;
      -webkit-tap-highlight-color: rgba(0, 0, 0, 0);
    }
    </style>
 
 
    <script>
    $(document).ready(function() {
 
      var owl = $("#owl-demo");
 
      owl.owlCarousel({
      autoPlay : 3000,
      items : 4, //10 items above 1000px browser width
      itemsDesktop : [1000,4], //5 items between 1000px and 901px
      itemsDesktopSmall : [900,3], // 3 items betweem 900px and 601px
      itemsTablet: [600,1], //2 items between 600 and 0;
      itemsMobile : false // itemsMobile disabled - inherit from itemsTablet option
     
      });
 
 
 
 
    });
    </script>
  <!-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------->
 
 
 
   <% end if %>
 
 
 
 
 
 
 
 
 
<!--#include file='alt.asp'-->
 
 
 
 
 
<script>

$(document).ready(function(){

           $(".thumb img").css({"cursor":"zoom-in"})
           $(".thumb img").mouseover(function(){
                       
           $('.devres').css('background-image', 'url("' + $(this).attr("src") + '")');
                       
                        
           
           })


})    




</script> 

                                     
 
          

    
    
    
    
</body>

</html>