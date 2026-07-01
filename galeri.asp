<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html>
<head>  
 
<META HTTP-EQUIV="Content-Type" content="text/html; charset=iso-8859-9"> 
   <meta name="viewport" content="width=device-width, initial-scale=1">

<title>Foto Galeri | Antalya Klima - Özüm Klima Mitsubishi Klima Antalya</title>   

<meta name="author" content="Mega Tasarım"/>   

<link rel="icon" type="image/png" href="fav.png" />   

<link href="css/font.css" rel="stylesheet" type="text/css"/>   

<script src="js/jquery-2.1.1.js"></script>  
<script src="js/jquery-ui.min.js"></script>               

   <link href="css/fontawesome-all.css" rel="stylesheet">



<link href="css/magnific-popup.css" rel="stylesheet" type="text/css">
 <script src="js/magnific-popup.js"></script>   
 
 <script src="js/darkbox.js"></script>   
 <link href="css/darkbox.css" rel="stylesheet" type="text/css">
 

    
</head>

<body> 

  
       
  
  
  
<!--  #include file='ust.asp' -->


<script>
$(document).ready(function() {
	$('.popup-with-form').magnificPopup({
		type: 'inline',
		preloader: false,
		focus: '#name',

		// When elemened is focused, some mobile browsers in some cases zoom in
		// It looks not nice, so we disable it:
		callbacks: {
			beforeOpen: function() {
				if($(window).width() < 700) {
					this.st.focus = false;
				} else {
					this.st.focus = '#name';
				}
			}
		}
	});
});


</script>
             
 

  
   <BR><BR>


    
     

<Div  class="tedaviler">    

 
 
<%     
set object = Server.CreateObJect("ADODB.RecordSet")
Sorgula = "Select * From galery WHERE gal_kat = 0 order by sira asc"
object.open Sorgula,baglanti,1,3   

	
Do while not object.Eof
%>  

   <a href="#"          data-darkbox="galeri/<%=object("foto")%>"             data-darkbox-group="one"   style="cursor: zoom-in;"> 
          <Div>  
                       <div class="TedResKap">
                         <div class="TedRes" style="background-image:url('galeri/<%=object("foto")%>')"></div>
                       </div>
          </Div>
          </a>    
     

	<%
object.MoveNext
Loop  
%>   
                             


</Div>            





         
     <BR>&nbsp;<BR>&nbsp;



       
      
            <script src="js/darkbox.js"></script>
      
      
     
      
      
      
      
      
      
      
    <!--  #include file='alt.asp' -->      

   

</body>
</html>