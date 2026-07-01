
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">  <html> 
<!--#include file='db.asp'-->

<head>
<meta name="viewport" content="width=device-width, initial-scale=1"> 
     
  <%   
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
  

%>


<title>Bul |  <%=ayars("firma") %> </title>    



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
<%=request("bul") %>                                                                                                                                                                                                                                                      
</div>

<div style="float:right">     


<select class="goturen"  >
<option value="#">Sırala</option>  
<option value="?g=<%=g%>&t=<%=t%>&s=<%=s%>&d=19"  <% if d=19 then %> selected <% end if %>> Ucuzdan Pahalıya</option>
<option value="?g=<%=g%>&t=<%=t%>&s=<%=s%>&d=91" <% if d=91 then %> selected <% end if %>>Pahalıdan Ucuza </option>
<option value="?g=<%=g%>&t=<%=t%>&s=<%=s%>&d=0" <% if d="0" then %> selected <% end if %>>Akıllı Sıralama </option>   
</select>
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
Sorgula = "Select * From products where isim LIKE '%"& request.form("bul") &"%' "& bul1 &"  "& bul2 &" and yayin=1 order by "&diz&"  "
object.open Sorgula,baglanti,1,3       


Do while not object.Eof     
%>      
      
      
      
      
      
                     <Div>   
                       
           
                                     
                                 <a href="<%=cevir(object("isim"))%>-<%=object("AffiliateID")%>">  
 <div class="urRes"  style="background-image:url('urunler/<%=object("foto1")%>');">
                                     <% if object("yeni")=1 then %><div class="yeni"><img  src="images/yeni.png"></div> <% end if %>
                                   <% if object("ozel5")=1 then %><div class="yeni"><img  src="images/indi.png"></div>   <% end if %>
                                     </div>  
                                                      
                                     
         <div  class="urAd" ><%=object("isim")%></div>   
                                        
                                            


     
<div class="detayal">İncele </div>    
            
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
%>                    
                       
                       

      </DIV>     
    
      
      </div>
     
<div style="clear:both"></div>
      
        
           </Div>  
           
           
           
           <BR><BR><BR>&nbsp;
 
<!--#include file='alt.asp'-->
 

    
</body>

</html>