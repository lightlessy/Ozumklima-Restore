
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">  <html> 

<head>
<meta name="viewport" content="width=device-width, initial-scale=1"> 
     
  <%   
 t= Trim(request.querystring("t"))
 s= Trim(request.querystring("s"))       
 d= Trim(request.querystring("d")) 
 bul = Trim(Request("bul"))
 bulSQL = Replace(bul,"'","''")

 if t<>"" then
   if IsNumeric(t)=false then t=""
 end if

 if s<>"" then
   if IsNumeric(s)=false then s=""
 end if
   
 if d="" then
   diz="sira asc"  
 elseif d="0" then    
   diz="sira desc"    
 elseif d="19" then    
   diz="fiyat asc"  
 elseif d="91" then    
   diz="fiyat desc"
 else
   diz="sira asc"
 end if       
  

%>


<title>Bul | Özüm Klima</title>    



 <META HTTP-EQUIV="Content-Type" content="text/html; charset=iso-8859-9"> 
  
    <link href='https://fonts.googleapis.com/css?family=Source+Sans+Pro:300' rel='stylesheet' type='text/css'>   
       
      <link rel="stylesheet" href="css/font.css">

      <link rel="shortcut icon" href="images/favicon.png" >      

   <link href="css/fontawesome-all.css" rel="stylesheet">
      <script src="js/jquery-2.1.1.js"></script>
      

</head>

<body>
       
             
       
<!--#include file='ust.asp'-->     

<BR><BR>


 
 <Div  class="navi" >

<div style="float:left">
<% if bul<>"" then %>
Arama: <%=Server.HTMLEncode(bul) %>
<% else %>
Arama
<% end if %>                                                                                                                                                                                                                                                      
</div>

<div style="float:right">     


<select class="goturen"  >
<option value="#">Sýrala</option>  
<option value="?bul=<%=Server.URLEncode(bul)%>&t=<%=t%>&s=<%=s%>&d=19"  <% if d="19" then %> selected <% end if %>> Ucuzdan Pahalýya</option>
<option value="?bul=<%=Server.URLEncode(bul)%>&t=<%=t%>&s=<%=s%>&d=91" <% if d="91" then %> selected <% end if %>>Pahalýdan Ucuza </option>
<option value="?bul=<%=Server.URLEncode(bul)%>&t=<%=t%>&s=<%=s%>&d=0" <% if d="0" then %>>Akýllý Sýralama </option>   
</select>
</div> 
 
<div style="clear:both;"></div>       



</Div>
 

     
  <script>
    $(document).ready(function(){

      $(".goturen").on('change', function () {
          var url = $(this).val();
          if (url && url !== "#") {
              window.location = url;
          }
          return false;
      });     
      
      
    });
</script>
 
    
     <Div style="width:100%; max-width:1200px; margin:0px auto; ">

    
 <BR><BR>



 <DIV   class="kutularic" >         
      
      
<%  

bul1=""
bul2=""
bulKosul=""

if t<>"" then
bul1="and tip="&t&""    
end if
           
if s<>"" then
bul2="and surface="&s&""    
end if

if bul<>"" then
bulKosul="and (isim LIKE '%"& bulSQL &"%' or descr LIKE '%"& bulSQL &"%' or keyw LIKE '%"& bulSQL &"%')"
end if

set object = Server.CreateObJect("ADODB.RecordSet")
Sorgula = "Select * From products where yayin=1 "& bulKosul &" "& bul1 &" "& bul2 &" order by "&diz&"  "
object.open Sorgula,baglanti,1,3       


if object.Eof then
%>

<div style="width:100%; padding:40px; text-align:center; font-size:1.2em; color:#555;">
<% if bul<>"" then %>
"<%=Server.HTMLEncode(bul)%>" icin sonuc bulunamadi.
<% else %>
Arama yapmak icin ustteki arama kutusunu kullanabilirsiniz.
<% end if %>
</div>

<%
else
Do while not object.Eof     
%>      
      
      
      
      
      
                     <Div>   
                       
           
                                     
                                 <a href="<%=cevir(object("isim"))%>-<%=object("AffiliateID")%>">  
 <div class="urRes"  style="background-image:url('urunler/<%=object("foto1")%>');">
                                     <% if object("yeni")=1 then %><div class="yeni"><img  src="images/yeni.png"></div> <% end if %>
                                   <% if object("ozel5")=1 then %><div class="yeni"><img  src="images/indi.png"></div>   <% end if %>
                                     </div>  
                                                      
                                     
         <div  class="urAd" ><%=object("isim")%></div>   
                                        
                                            


     
<div class="detayal">Ýncele </div>    
            
<div class="yildiz">
<%  
set yorumcu = Server.CreateObJect("ADODB.RecordSet")
Sorgula = "Select * From yorumlar where urun ="& object("AffiliateID")&" and rating > 3 and onay =1"
yorumcu.open Sorgula,baglanti,1,3 
if not yorumcu.eof then   
%>
<% for i=1 to  yorumcu("rating") %>
<i class="fa fa-star" aria-hidden="true"  style="font-size:.7em; color:#F39F4B" ></i>    
<% next %>
<span style="font-weight:normal; font-size:.6em; letter-spacing:0px"> Yorum (<%=yorumcu.recordcount %>) </span>  
<% end if %>    
</div>
 
  </a>   

</Div>
                       
                       
           
           <%  
object.MoveNExt
Loop
end if
%>                    
                       
                       

      </DIV>     
    
      
      </div>
     
<div style="clear:both"></div>
      
        
           </Div>  
           
           
           
           <BR><BR><BR>&nbsp;
 
<!--#include file='alt.asp'-->
 

    
</body>

</html>