<html>
 <head>
  <title>PANEL </title> 
  <meta http-equiv="Content-Type" content="text/html; charset=utf-8">        
<link rel=stylesheet type="text/css" href="font.css"> 

<% grup = request.querystring("grup")
t = request.querystring("t")%>

</head>     
   
   
  <!--#include file="menu.asp"-->   
    
  <div class="baslik">Ürün Düzenle </div>
  
   
   <link rel="stylesheet" href="../css/jquery-ui.css">
  <style>
  #sortable { list-style-type: none; margin: 0; padding: 0; font-family:Open Sans;font-size:13px; }    
    #sortable li {  font-family:open sans; color:#464646; font-size:0.9em; max-width:700px; cursor:move;  }
  #sortable div {display:inline-block; vertical-align:top;  padding:10px 5px 5px 5px;}
  
  </style>

  <script src="../js/jquery-ui.js"></script>
  <script>
  $( function() {
    $( "#sortable" ).sortable();
    $( "#sortable" ).disableSelection();
  } );
  </script>    
  
  
  
     <BR><BR>
  
<%


if Request.Querystring("psil2")=1 then 


Set Rs = Server.CreateObject("Adodb.Recordset")
rSQL = "Select * from products  where AffiliateID = "& Request("id")
Rs.Open rSQL, baglanti, 1, 3 

Rs.Delete 
Rs.Update 

rs.close
set rs = nothing 
Response.Redirect Request.ServerVariables("HTTP_REFERER")         

end if      %>   

 



  



<table align=center border=0><tr><td align=center>
<%

Set mods1 = Server.CreateObject("Adodb.Recordset")
rSQL = "Select *  from  anagrup "
mods1.Open rSQL, baglanti, 1, 3    

Do while not mods1.Eof   %>  



<div style="display:block; border:1px solid #DADADA; background-color:#FFFFFF;  width:250px;padding:8px; margin:5px" class="font"><a href="urunler.asp?grup=<%=mods1("id")%>"><% if mods1("id")=Cint(grup) then %><font color="#AE0000"><b><% end if %>  <%=mods1("isim")%> </b></font></a>

 <%    
 if Cint(grup) = mods1("id")  then
 
Set mods = Server.CreateObject("Adodb.Recordset")
rSQL = "Select * From tip where grup ="& mods1("id") &"  order by sira desc"
mods.Open rSQL, baglanti, 1, 3       

Do while not mods.Eof         
%>
                      
<div style="margin-top:5px"><a href="urunler.asp?grup=<%=mods1("id")%>&t=<%=mods("id")%>#g"><%=mods("isim")%></a></div>

<%  
mods.MoveNExt
Loop  
end if
%>   

</div>

  
<%     
mods1.MoveNExt
Loop   

%>  

   
<a name="a"></a>

<%

if grup <>"" then   

if t<>"" then
bul="and tip= "&t&""     

Set gpst = Server.CreateObject("Adodb.Recordset")
rSQL = "Select *  from  tip where id= "&t &" "
gpst.Open rSQL, baglanti, 1, 3    

end if

Set mods = Server.CreateObject("Adodb.Recordset")
rSQL = "Select * from products  where anagrup="&grup &"  "& bul &" and yayin=1 order by sira asc"
mods.Open rSQL, baglanti, 1, 3

Set gps = Server.CreateObject("Adodb.Recordset")
rSQL = "Select *  from  anagrup where id= "&grup &" "
gps.Open rSQL, baglanti, 1, 3  

      

%>  


<BR><hr style="height: 1px;background-color: #AEAEAE; border:none;">      <BR> 

   <form action="urun-ana.asp" method="post"> 
   
<a name="g"></a>

 <ul id="sortable">  
           
<font color="#AE0000" size=4><b><%=gps("isim")%></b> <% if t<>"" then %> / <%=gpst("isim")%> <% end if %></font><BR><BR> <BR>
 

     <font class="font" style="color:#4E4E4E; font-size:12px"> SIRALAMAYI DEĞİŞTİRMEK İÇİN SÜRÜKLEYİP BIRAKINIZ.       </font>        
        <BR><BR>


      
<%    k=1 
 Do while not mods.Eof      %>   
  
     
    <li class="ui-state-default" id="<%=mods("AffiliateID")%>">    

<div style="text-align:left; width:500px"><%=mods("isim")%> &nbsp;- &nbsp; Stok: <%=mods("stok")%> </div>

<div>
<input type="checkbox" name="Veri<%=k%>" value="<%=mods("AffiliateID")%>" <% if mods("urunana")=1 then %>checked <% end if %>> <font size=1>Ana Sayfada Göster </font>
</div>
<div> 
<a href="edits.asp?id=<%=mods("AffiliateID")%>&g=<%=mods("anagrup")%>&tip=<%=mods("tip")%>"><img src="images/edit.png" width=20></a>
<a href="?psil2=1&id=<%=mods("AffiliateID")%>" onclick="return confirm('Kayıt silinecek. Onaylıyor Musunuz?')"><img src="images/del.png" width=20></a>
</div>

   </li>


<%    
  k = k + 1 
mods.MoveNExt
Loop
%>      

         </ul>


            <input type="submit" value="Seçili Ürünleri Ana Sayfada Göster"><input type="hidden" name="toplam" value="<%=Int(k-1)%>"></td> 

                  </form>       
         
           <script>

$(document).ready(function () {
    $('ul').sortable({   
    
    
    
        axis: 'y',
        stop: function (event, ui) { 
                                       

        
        
	        var sira = ui.item.index()     
	    var id = ui.item.attr("id");       
               
var siralama = [];          
          
$( "#sortable li" ).each(function( index ) {     

siralama.push($(this).attr("id")+"-"+index);


});
          

	        
	                 
  $.ajax({
       type: 'POST',
       url: 'sirala2.asp',
       data: "giden="+siralama,
       success: function(ajaxCevap) {
        

$(".sonuc").text(ajaxCevap);
                 }

            });


            
	}   
	
	
	
    });
});   
   
   </script>

       <% end if %>
                        
       
         
</td></tr></table>
                         
<BR><BR>
&nbsp;

