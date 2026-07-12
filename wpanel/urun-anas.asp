<%@ Language="VBScript" CodePage="65001" %>
<%
Response.Buffer = True
Response.CodePage = 65001
Response.Charset = "utf-8"
%>
<html>
 <head>
  <title>PANEL </title>
  <meta charset="utf-8">
<link rel=stylesheet type="text/css" href="font.css">

<% grup = request.querystring("grup")%>




</head>


  <!--#include file="menu.asp"-->


  <div class="baslik"> Ana Sayfa Ürünleri </div>

   <link rel="stylesheet" href="../css/jquery-ui.css">
  <style>
  #sortable { list-style-type: none; margin: 0; padding: 0;  }
    #sortable li {  font-family:open sans; color:#464646; font-size:0.9em; max-width:600px; cursor:move;  }
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


   <table align=center><tr><td>


     <font class="font" style="color:#4E4E4E; font-size:12px"> SIRALAMAYI DEĞİŞTİRMEK İÇİN SÜRÜKLEYİP BIRAKINIZ.       </font>
        <BR><BR>



 <ul id="sortable">
<%

Set mods = Server.CreateObject("Adodb.Recordset")
rSQL = "Select * from products  where urunana=1 order by anas asc"
mods.Open rSQL, baglanti, 1, 3



  Do while not mods.Eof      %>



  <li class="ui-state-default" id="<%=mods("AffiliateID")%>">


 <div style="width:400px; text-align:left; "><%=mods("isim")%></div>

 <div><a href="urun-ana.asp?kald=1&rid=<%=mods("AffiliateID")%>" title="Ana Sayfadan Kaldır"><img src="images/rem.png" width=20> </a></div>

 <div>
<a href="edits.asp?id=<%=mods("AffiliateID")%>&g=<%=mods("anagrup")%>&tip=<%=mods("tip")%>"><img src="images/edit.png" width=20></a>
</div>

   </li>


<%

mods.MoveNExt
Loop
%>

         </ul>



  </td></tr></table>


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
       url: 'sirala.asp',
       data: "giden="+siralama,
       success: function(ajaxCevap) {


$(".sonuc").text(ajaxCevap);
                 }

            });



	}



    });
});

   </script>

<BR><BR>
&nbsp;

