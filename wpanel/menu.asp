<html>
 <head>
  <title>PANEL </title> 
  <meta name="viewport" content="width=device-width, initial-scale=1">         

  <meta http-equiv="Content-Type" content="text/html; charset=windows-1254">        
<link rel=stylesheet type="text/css" href="font.css"> 

    
<%
if  Request.Cookies("ozum")("wpAfLogin") <> "true" then
	response.redirect "index.asp"       
		

	end if
	%>
        

</head>     

<!--#INCLUDE file="dbase.asp"-->      

<%
Set ayars = Server.CreateObject("ADODB.Recordset")
sqly = "SELECT * FROM ayarlar"
ayars.open sqly, baglanti,1,3    
%>

  <script src="../js/jquery-2.1.1.js"></script>   

       
<link rel="stylesheet" type="text/css" href="style.css">

<div class="container">

 


<a class="toggleMenu" href="#"  style="font-size:x-large">Menu</a>
<ul class="nav">  

	<li  class="test">
		<a href="#">KATEGORİLER</a> 
		
		<ul>
<li><a href="mainsil.asp">Ana Kategoriler</a>   </li>
<li><a href="tipsil.asp">Alt Kategoriler</a>      </li>
<li><a href="cargo-edit.asp">Kargo Düzenleme</a>     </li>
  </ul>
		
	</li>       
	              
	         
	         
			<li  class="test">
		<a href="#">ÜRÜNLER</a> 
		
		<ul>
		
<li><a href="join.asp">Ekle</a>      </li>
<li><a href="urunler.asp">Düzenle</a>   </li>       
<li><a href="indirim-urunler.asp">İndirimli Ürünler</a>      </li> 
<li><a href="pasif-urunler.asp">Pasif Ürünler</a>      </li> 
<li><a href="urun_bul.asp">Ürün Bul</a>       </li>     
<li><a href="urun-hitler.asp">Popüler</a>   </li>
  </ul>
		
	</li>        
	
	
	
	
	
				<li  class="test">
		<a href="#">SİPARİŞLER</a> 
		
		<ul>

 <li> <a href="orders.asp">Devam Eden </a>     </li>    
 <li> <a href="ar_orders.asp">Kargolanan</a>       </li>
 <li> <a href="iptal_orders.asp">İptal Olan</a>     </li> 
 
  </ul>
		
	</li>  
	
	
	
	
	
	
	
	
		<li  class="test">
		<a href="#">MÜŞTERİLER</a> 
		
		<ul>
<li><a href="uye-list.asp">Müşteri Listesi</a>       </li>  
<li><a href="yorumlar.asp">Yorumlar</a>       </li>

  </ul>
		
	</li>  		
	      
	
	
	

	
		<li  class="test">
		<a href="#">SLAYT</a> 
		
		<ul>
<li><a href="slayt.asp">Ekle</a>   </li>
<li><a href="slayt_sil.asp">Düzenle</a>    </li>
  </ul>
		
	</li>  
	
     
   
   		<li  class="test">
		<a href="#">İÇERİK</a> 
		
		<ul>
<li><a href="sayfa-ekle.asp">Ekle</a>   </li>
<li><a href="sayfa-sil.asp?id=1">Kurumsal Düzenle</a>    </li>        
<li><a href="sayfa-sil.asp?id=2">Hizmetler Düzenle</a>    </li>  
<li><a href="sayfa-sil.asp?id=3">Blog Düzenle</a>    </li>  
<li><a href="refer.asp">Referans Ekle</a>   </li>
<li><a href="refer_sil.asp">Referans Düzenle</a>    </li>
  </ul>
		
	</li>  
	
	
	
			<li  class="test">
		<a href="#">GALERİ</a> 
		
		<ul>
<li><a href="add_gal_foto.asp">Fotoğraf Ekle</a>   </li>
<li><a href="galeri_edit.asp">Fotoğraf Düzenle</a>    </li>
  </ul>
		
	</li>  
	
	   
   		<li  class="test">
		<a href="#">AYARLAR</a> 
		
		<ul>      
<li><a href="ayar.asp">İletişim Bilgileri</a>    </li>
<li><a href="ayar.asp">Şifre Değişimi</a>   </li>
<li><a href="ayar.asp">Banka Bilgileri</a>    </li>
<li><a href="ayar.asp">Sosyal Medya</a>    </li>
  </ul>
		
	</li>  



	
	 	<li><a href="exit.asp" onClick='javascript:window.close()'>ÇIKIŞ  </a></li>  
	
	</ul>
</div>
	
	


             <script type="text/javascript" src="script.js"></script>
          <script>
 
 $(document).ready(function(){                                                                                                            
              
                 function mobilse() {       }                                            
              
              
if(  ( /Android|webOS|iPhone|iPad|iPod|BlackBerry|IEMobile|Opera Mini/i.test(navigator.userAgent))      ) {   mobilse()   }   

else    { 

                $("body").css({"font-size":"15px"}) 

      }
 
 
 })
 
 
 </script> 
	          
