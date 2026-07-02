  <div class="whatweb">
  <a href="https://api.whatsapp.com/send?phone=9<%=replace(ayars("gsm")," ","")%>&text=https://www.ozumklima.com" target=_Blank style="color:#4D4D4D;" onclick="analyticsTrack('whatsapp_click', {location: 'footer', method: 'api.whatsapp.com'});"> <img src="images/what2.png" width=60> <BR>WhatsApp Destek   </a>    
</div>



    
<DIV   style="background-color:#353535" >
     <Div  class="alt"  >           
        
        
        
            <Div style="float:left"  class="altmenu"   >
                       
                       <Div>
                                <a href="klima.asp?t=36">      <span> Ev Ofis Klimalar� </span>    </a>

          
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
                             <a href="havalandirma.asp">       <span> Havaland�rma Sistemleri  </span> </a> 

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
                       
                       
   
            
            </Div>
                       
              
              
              
            <Div style="float:right; max-width:200px"  class="bizeulasin" >
                   <font style="font-size:1.2em">Bize Ula��n<BR></font>
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
                   
                  

                   <Div>� 2025 <a href="http://www.ozumlkila.com" title="Antalya Klima">�z�m Klima </a> - T�m Haklar� Sakl�d�r. &nbsp;  | &nbsp;  <a  href="http://www.megatasarim.com"   target="_blank" >Mega Tasar�m </a></Div>
       
                   

</Div>
</DIV>     
                      


<!--#include file='analytics.asp'-->
             

