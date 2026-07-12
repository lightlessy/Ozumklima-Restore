  <div class="whatweb">
  <a href="https://api.whatsapp.com/send?phone=9<%=replace(ayars("gsm")," ","")%>&text=https://www.ozumklima.com" target=_Blank style="color:#4D4D4D;" onclick="analyticsTrack('whatsapp_click', {location: 'footer', method: 'api.whatsapp.com'});"> <img src="images/what2.png" width=60> <BR>WhatsApp Destek   </a>    
</div>



    
<DIV   style="background-color:#353535" >
     <Div  class="alt"  >           
        
        
        
            <Div style="float:left"  class="altmenu"   >
                       
                       <Div>
                                <a href="klima.asp?t=36">      <span> Ev Ofis Klimalar&#305; </span>    </a>

          
 <%  

set alturun = Server.CreateObJect("ADODB.RecordSet")
Sorgula = "Select top 9 * From products where tip = 36 order by sira desc"
alturun.open Sorgula,baglanti,1,3  

Do while not alturun.Eof                                                                                                                                                                                                                  
%>   
  
<a href="<%=cevir(alturun("isim"))%>-<%=alturun("AffiliateID")%>"><%=Left(alturun("isim"),30) %> </a>   <BR>
<%  
alturun.MoveNExt
Loop
%>                       



                       
            
                       </Div>     
                       
                       
                       <Div>
                            <a href="klima.asp?t=37">    <span> Ticari Tip Klimalar </span>      </a>

 <%  

set alturun = Server.CreateObJect("ADODB.RecordSet")
Sorgula = "Select top 9 * From products where tip = 37 order by sira desc"
alturun.open Sorgula,baglanti,1,3  

Do while not alturun.Eof                                                                                                                                                                                                                  
%>   
  
<a href="<%=cevir(alturun("isim"))%>-<%=alturun("AffiliateID")%>"><%=Left(alturun("isim"),30) %> </a>   <BR>
<%  
alturun.MoveNExt
Loop
%>    
                 
            
                       </Div>
                       
                       
                       
                       <Div>
                               <a href="vrf_sistemler.asp">    <span> VRF Merkezi Sistemler </span>   </a>
 <%  

set alturun = Server.CreateObJect("ADODB.RecordSet")
Sorgula = "Select top 9 * From products where tip = 41 order by sira desc"
alturun.open Sorgula,baglanti,1,3  

Do while not alturun.Eof                                                                                                                                                                                                                  
%>   
  
<a href="<%=cevir(alturun("isim"))%>-<%=alturun("AffiliateID")%>"><%=Left(alturun("isim"),30) %> </a>   <BR>
<%  
alturun.MoveNExt
Loop
%>    

                       </Div>
                       
                       
                       
                       <Div>
                             <a href="havalandirma.asp">       <span> Havaland&#305;rma Sistemleri  </span> </a> 

 <%  

set alturun = Server.CreateObJect("ADODB.RecordSet")
Sorgula = "Select top 9 * From products where tip = 39 order by sira desc"
alturun.open Sorgula,baglanti,1,3  

Do while not alturun.Eof                                                                                                                                                                                                                  
%>   
  
<a href="<%=cevir(alturun("isim"))%>-<%=alturun("AffiliateID")%>"><%=Left(alturun("isim"),30) %> </a>   <BR>
<%  
alturun.MoveNExt
Loop
%>    
            
                       </Div>  

                       <Div>
                             <a href="h.asp?id=12">       <span> Antalya Hizmetleri </span> </a> 
<a href="antalya-klima-servisi.asp">Antalya Klima Servisi</a>   <BR>
<a href="dosemealti-klima-servisi.asp">Dosemealti Klima Servisi</a>   <BR>
<a href="antalya-klima-montaji.asp">Antalya Klima Montaji</a>   <BR>
<a href="antalya-klima-bakimi.asp">Antalya Klima Bakimi</a>   <BR>
<a href="antalya-mitsubishi-electric-klima.asp">Mitsubishi Electric Klima</a>   <BR>
<a href="antalya-vrf-sistemleri.asp">Antalya VRF Sistemleri</a>   <BR>
                       </Div>

                       <Div>
                             <a href="klima-rehberi.asp">       <span> Klima Rehberi </span> </a> 
<a href="40-m2-oda-kac-btu-klima.asp">40 m2 Oda Kac BTU?</a>   <BR>
<a href="vrf-mi-multi-split-mi.asp">VRF mi Multi Split mi?</a>   <BR>
<a href="klima-kesfi-nasil-yapilir.asp">Klima Kesfi Nasil Yapilir?</a>   <BR>
<a href="antalya-villa-klima-cozumleri.asp">Villa Klima Cozumleri</a>   <BR>
                       </Div>
                       
                       
   
            
            </Div>
                       
              
              
              
            <Div style="float:right; max-width:200px"  class="bizeulasin" >
                   <font style="font-size:1.2em">Bize Ula&#351;&#305;n<BR></font>
                  <span class="alttel"><i class="fas fa-phone"></i>       <%=ayars("tel") %> <BR><BR></span>
                  
                   <B>Adres: </B><BR>
              <%=ayars("adres") %><BR><BR>


                  <B>E-Posta:</B>      <BR>
                 <a href="mailto:<%=ayars("eposta") %>" target=_Blank style="color:#DBF2EB"> <%=ayars("eposta") %></a>  <BR>
            
            
            </Div>              
              
              
                    
              
              
        
        <div style="clear:both" ></div>
        </Div>
</DIV>      
      


      
<DIV  style="background-color:#353535; border-top:1px solid #7E7E7E "  >
<Div  class="enalt" >     
                   
                  

                   <Div>&copy; 2025 <a href="https://www.ozumklima.com" title="Antalya Klima">&Ouml;z&uuml;m Klima </a> - T&uuml;m Haklar&#305; Sakl&#305;d&#305;r. &nbsp;  | &nbsp;  <a  href="http://www.megatasarim.com"   target="_blank" >Mega Tasar&#305;m </a></Div>
       
                   

</Div>
</DIV>     
                      

<link href="css/home-polish.css" rel="stylesheet" type="text/css">
<!--#include file='js/home-polish.js'-->

<!--#include file='analytics.asp'-->