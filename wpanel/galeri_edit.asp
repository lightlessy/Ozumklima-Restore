   
<!--#INCLUDE file="menu.asp"-->        

    
   <link rel="stylesheet" href="../css/jquery-ui.css">
  <style>
  #sortable { list-style-type: none; margin: 0; padding: 0; width: 1100px; }
  #sortable li { display:inline-block; margin: 0 20px 20px 20px; padding: 0.4em; padding-left: 0.8em; font-size: 1.4em;  }
  </style>
  <script src="../js/jquery-2.1.1.js"></script>
  <script src="../js/jquery-ui.js"></script>
  <script>
  $( function() {
    $( "#sortable" ).sortable();
    $( "#sortable" ).disableSelection();
  } );
  </script>     
  
  

  
  

<body> 
              
   
       <div class="baslik">Galeri Düzenle</div>     
   
<BR><BR><center>            

                                  

<%if request("delete")<>"" then  %>

<%


Set Rs = Server.CreateObject("Adodb.Recordset")
rSQL = "Select * from galery where id = "& Request.Form("id")
Rs.Open rSQL, baglanti, 1, 3 
Rs.Delete ' Sil
Rs.Update ' Güncelle

rs.close
set rs = nothing 

Response.Redirect Request.ServerVariables("HTTP_REFERER")
%>        
<% end if %>      




<%if request("update")="1" then  


id=Request("id")
isim=Request("isim")
%>

<%
a=split(id,",") 
b=split(isim,",") %>

<%uzunluk=uBound(a)     %>

<% for i=0 to uzunluk %>

<%
Set Rs = Server.CreateObject("Adodb.Recordset")
rSQL = "Select * from galery where  id = "&a(i)
Rs.Open rSQL, baglanti, 1, 3 
Rs("isim")=replace(b(i)," ","")
Rs.Update     

next   

Response.Redirect "galeri_edit.asp?gal_kat="&Request("gal_kat")&"&altkat="&Request("altkat")&""

end if %>    


 

<%if request("foto")<>"" then  %>

<%

Set Rs = Server.CreateObject("Adodb.Recordset")
rSQL = "Select * from gal_kategoriler where id = "& Request("gal_kat")
Rs.Open rSQL, baglanti, 1, 3 

Rs("foto") = Request("foto")  
Rs.Update  


rs.close
set rs = nothing 

Response.Redirect "galeri_edit.asp?gal_kat="&Request("gal_kat")&"&altkat="&Request("altkat")&""
%>        
<% end if %>   


 
<font class="font" style="color:#E10000; font-size:15px"> SIRALAMAYI DEÐÝÞTÝRMEK ÝÇÝN RESMÝ SÜRÜKLEYÝP BIRAKINIZ.       </font>
 <BR><BR>
                           
 
 <FORM NAME="MyForm" METHOD="POST"  	ACTION="galeri_edit.asp"> 

         <INPUT TYPE="hidden"  NAME="gal_kat" value="<%=request("gal_kat")%>"> 
               <INPUT TYPE="hidden"  NAME="altkat" value="<%=request("altkat")%>">  
               
<ul id="sortable">            


 
<%     
set object = Server.CreateObJect("ADODB.RecordSet")
Sorgula = "Select * From galery WHERE gal_kat = 0 order by sira asc"
object.open Sorgula,baglanti,1,3   
if not object.eof then

 	i = 1
	If (i mod 5)=0 Then 
      StrRow="5"                                                                     
	Else 
	StrRow=(i mod 5)        
	End If            
	
	
	
Do while not object.Eof
%>    


  <li class="ui-state-default" style="display:inline-block; " id="<%=object("id")%>">    
  

   
      <INPUT TYPE="hidden"  NAME="id" value="<%=object("id")%>">    <BR>
       
<a href="../galeri/<%=object("foto")%>" target=_Blank><img src="../galeri/<%=object("foto")%>"  border=0 width=200 height=150></a>
 <BR>   

 <a href="imageupdate_galeri.asp?id=<%=object("id")%>&gal_kat=<%=object("gal_kat")%>" title="Foto Güncelle"><img src="images/guncel.png" width=25></a>
&nbsp;&nbsp;
 <a href="sil_galfoto.asp?id=<%=object("id")%>&gal_kat=<%=object("gal_kat")%>&altkat=<%=request("altkat")%>" title="Foto Sil" onclick="return confirm('Kayýt silinecek. Onaylýyor Musunuz?')"><img src="images/del.png" width=25></a>

   </li>
	<%
object.MoveNext
Loop  
end if	 %>     

      </ul>
    
</div>           
   
      <INPUT TYPE="hidden"  NAME="update" value="1">         
   
  	<INPUT TYPE=SUBMIT VALUE="KAYDET">  
   
   
   
   
        
   <script>

$(document).ready(function () {
    $('ul').sortable({   
    
    
    
        axis: 'x,y',
        stop: function (event, ui) { 
                                       

        
        
	        var sira = ui.item.index()     
	    var id = ui.item.attr("id");       
               
var siralama = [];          
          
$( "#sortable li" ).each(function( index ) {     

siralama.push($(this).attr("id")+"-"+index);


});
          

	        
	                 
  $.ajax({
       type: 'POST',
       url: 'sirala.asp',
       data: "gal=<%=Request("gal_kat")%>&giden="+siralama,
       success: function(ajaxCevap) {
        

$(".sonuc").text(ajaxCevap);
                 }

            });


            
	}   
	
	
	
    });
});   
   
   </script>
               

   

 
</body>
</html>


                                
      <BR><BR><BR>
     
 