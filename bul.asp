
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
<option value="?bul=<%=Server.URLEncode(bul)%>&t=<%=t%>&s=<%=s%>&d=0" <% if d="0" then %> selected <% end if %>>Akýllý Sýralama </option>   
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
urunIsim = object("isim") & ""
urunID = object("AffiliateID") & ""
urunFoto = object("foto1") & ""
urunLink = "#"

if urunID<>"" then
  if urunIsim<>"" then
    urunLink = cevir(urunIsim) & "-" & urunID
  else
    urunLink = "detay.asp?" & urunID
  end if
end if
%>      
      
      
      
      
      
                     <Div>   
                       
           
                                     
                                 <a href="<%=urunLink%>">  
 <% if urunFoto<>"" then %>
 <div class="urRes"  style="background-image:url('urunler/<%=urunFoto%>');"></div>  
 <% else %>
 <div class="urRes"></div>
 <% end if %>
                                                      
                                     
         <div  class="urAd" ><%=Server.HTMLEncode(urunIsim)%></div>   
                                        
                                            


     
<div class="detayal">Ýncele </div>    
            
 
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