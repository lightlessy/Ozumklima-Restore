
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


 <META HTTP-EQUIV="Content-Type" content="text/html; charset=iso-8859-9"> 
  
    <link href='https://fonts.googleapis.com/css?family=Source+Sans+Pro:300' rel='stylesheet' type='text/css'>   
       
      <link rel="stylesheet" href="css/font.css">

      <link rel="shortcut icon" href="images/favicon.png" >      

   <link href="css/fontawesome-all.css" rel="stylesheet">
      <script src="js/jquery-2.1.1.js"></script>
      

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

     
<div class="detayal">�ncele </div>    
            

 
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