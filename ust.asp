<%
Session.CodePage = 1254
Session.LCID = 1055
%>
<%Response.Charset="ISO-8859-9"
Response.Charset="Windows-1254"
response.ContentType="text/HTML"
%>


<!--#include file='db.asp'-->
 



 <Div class="menuLogoTas"  >
 

  <div id="nav-icon2" class="ackapa" >
  <span></span>
  <span></span>
  <span></span>
  <span></span>
  <span></span>
  <span></span>
</div>

<a href="tel:0 242 229 71 39"><div class="fon"><i class="fas fa-phone-square"></i></div></a>
   
  
  
         <Div  class="umenu"  ">     
 
  <div class="logotas">   <a href="https://ozumklima.com">  <img  src="images/logo.png"  width=100%></a>   </div>
     
  
                
                    <div class="ustsosyal">          
            <a href="<%=ayars("face")%>" target=_blank><div><i class="fab fa-facebook-f"></i> </div></a>
         <a href="<%=ayars("insta")%>" target=_blank><div><i class="fab fa-instagram"></i> </div></a>        
         </div>   
         
     <div class="AramaYap">  
             <FORM METHOD="POST" action="bul.asp" name="aform" >  
                   <input type="text" name="bul" required    class="searchTerm"  placeholder="Ne aramýþtýnýz?" style="outline:none; border:none;; border-bottom:1px solid #EAEAEA; width:80%">  
                     <span class="buyutec"> <input type="submit" value=""  class="searchButton " ></span>
                     </form> 

           </Div>
  
   

</Div>
          
       
                <div style="clear:both;"></div>  
       
  <div class="serit">
  
         <Div  class="menu"  ">     
              <a href="https://ozumklima.com"><div>ANA SAYFA</div></a>   
        <a href="kurumsal.asp"><div>KURUMSAL</div></a>
    <a href="h.asp?id=12"><div>HÝZMETLER</div></a>
    <a href="klima.asp"><div>KLÝMA</div></a>     
             <a href="havalandirma.asp"><div>HAVALANDIRMA SÝSTEMLERÝ</div></a>     
         <a href="vrf_sistemler.asp"><div>VRF SÝSTEM</div></a>     
                 <a href="mitsubishi-electric-isi-pompasi-(heat-pump)-165"><div>ISI POMPASI </div></a>               
               <a href="galeri.asp"><div>GALERÝ</div></a>
            <a href="referanslar.asp"  ><div>  REFERANSLAR</div></a>   
    <a href="iletisim.asp"><div>ÝLETÝÞÝM </div></a> 
   
    </Div>
  </div>
  
                                

  
  

 <style>    
 
 #nav-icon2 {
  width: 30px;
  height: 35px;       
  margin-right:5px; 
  margin-top:3px;
z-index:999;
-webkit-transform: rotate(0deg);
  -moz-transform: rotate(0deg);
  -o-transform: rotate(0deg);
  transform: rotate(0deg);
  -webkit-transition: .5s ease-in-out;
  -moz-transition: .5s ease-in-out;
  -o-transition: .5s ease-in-out;
  transition: .5s ease-in-out;
  cursor: pointer;
}

 
 
 #nav-icon2 {
}

#nav-icon2 span {
  display: block;
  position: absolute; 

  height: 4px;
  width: 50%;
  background: #3B3B3B;
  opacity: 1; 
 top:0px;
  -webkit-transform: rotate(0deg);
  -moz-transform: rotate(0deg);
  -o-transform: rotate(0deg);
  transform: rotate(0deg);
  -webkit-transition: .25s ease-in-out;
  -moz-transition: .25s ease-in-out;
  -o-transition: .25s ease-in-out;
  transition: .25s ease-in-out;
}

#nav-icon2 span:nth-child(even) {
  left: 50%;
  border-radius: 0 9px 9px 0;
}

#nav-icon2 span:nth-child(odd) {
  left:0px;
  border-radius: 9px 0 0 9px;
}

#nav-icon2 span:nth-child(1), #nav-icon2 span:nth-child(2) {
  top: 0px;
}

#nav-icon2 span:nth-child(3), #nav-icon2 span:nth-child(4) {
  top: 9px;
}

#nav-icon2 span:nth-child(5), #nav-icon2 span:nth-child(6) {
  top: 18px;
}

#nav-icon2.open span:nth-child(1),#nav-icon2.open span:nth-child(6) {
  -webkit-transform: rotate(45deg);
  -moz-transform: rotate(45deg);
  -o-transform: rotate(45deg);
  transform: rotate(45deg);
}

#nav-icon2.open span:nth-child(2),#nav-icon2.open span:nth-child(5) {
  -webkit-transform: rotate(-45deg);
  -moz-transform: rotate(-45deg);
  -o-transform: rotate(-45deg);
  transform: rotate(-45deg);
}

#nav-icon2.open span:nth-child(1) {
  left: 5px;
  top: 7px;
}

#nav-icon2.open span:nth-child(2) {
  left: calc(50% -7px);
  top: 7px;
}

#nav-icon2.open span:nth-child(3) {
  left: -50%;
  opacity: 0;
}

#nav-icon2.open span:nth-child(4) {
  left: 100%;
  opacity: 0;
}

#nav-icon2.open span:nth-child(5) {
  left: 4px;
  top: 18px;
}

#nav-icon2.open span:nth-child(6) {
  left: calc(50% - 0px);
  top: 18px;
}   

 
 </style>



     <script>

$(document).ready(function(){   
         var acikmi=false;   
       
		$('#nav-icon2').click(function(){  
		          	$(this).toggleClass('open');    
		          	acikmi=!acikmi
                               if (acikmi ) {  $(".Menu").slideDown()  }    else {  $(".Menu").slideUp() }  
		 	});      
		   
	        
});

</script>  


  
           
                     
  
  <script>         
  $(document).ready(function(){             
  
  
   
                         
   var ekranGen=$(window).width();
            
            
            if (  ekranGen > 1024 ){  
   
  
  

         var didScroll = false;

window.onscroll = doThisStuffOnScroll;

function doThisStuffOnScroll() {     didScroll = true;    }
           

             /*  $(".genislik").css({ width: "300px" },200) ; */

setInterval(function() {  


    if(didScroll) {
                     didScroll = false;
                     var scroll = $(window).scrollTop();
                     if (scroll > 80) {    
                                                 $(".logo").attr("src","images/logo2.png")
                                                $(".logo").stop().animate({ width: "200px" },200) ;        
                            
                                    
                                               $(".menuLogoTas").stop().animate({ paddingTop: "0px" },200) ;
                                              $(".genislik").stop().animate({ width: "250px" },200) ;    
                                               
                                             }
                                                     
                                     else {    
                                                $(".logo").attr("src","images/logo.png") 
                                                 $(".logo").stop().animate({ width: "250px" },200) ;
                                   
                                  
                                             $(".menuLogoTas").stop().animate({ paddingTop: "0px" },200) ;
                                              $(".genislik").stop().animate({ width: "300px" },200) ; 
                                              
                                               }
                }    
    
    
    
}, 100);

                

         
         
         }       else {   $(".logo").attr("src","images/logo2.png")     }
         

             
         

         
         
          var acikmi=false;
          
   

    
             
     
        $( ".ackapa" ).click(function(){      
        
            
        

            acikmi=!acikmi
            if (acikmi ) {  
                                $(".menu").slideDown() ; 
                                $( ".acKapa" ).css({"opacity":".4"})
                            }    
            
            else          {      
                             
                             $(".menu").slideUp() ;
                             $( ".ackapa" ).css({"opacity":"1"})
                            
                           }  
     
         })
   
         
               
         
           
         
         
                                   if (  ekranGen < 1024 ){ 
                                   
                                    $( ".MenuVeicerik" ).before().click(function(){
             
                      acikmi=!acikmi
            if (acikmi ) {  
                                $(".solMenu").slideDown() ; 
                                 
                            }    
            
            else          {      
                             
                             $(".solMenu").slideUp() ;
                            
                            
                           }  
     
         })
                                   
                                    }
     

         

   })
  </script>
