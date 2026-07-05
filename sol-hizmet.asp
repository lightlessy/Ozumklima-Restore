   <span class="solMenuBas" >HÝZMETLER</span>
     
   <%  

set object = Server.CreateObJect("ADODB.RecordSet")
Sorgula = "Select * From sayfalar where ana=2 order by sira asc"
object.open Sorgula,baglanti,1,3  

Do while not object.Eof                                                                                                                                                                                                                  
%>   

 <a href="h.asp?h=<%=cevir(object("isim")) %>&id=<%=object("id") %>"  ><div><i class="fas fa-chevron-right"></i> <%=object("isim") %></div></a>   
 
<%  
object.MoveNExt
Loop
%>

<span class="solMenuBas" style="margin-top:18px; display:block;" >ANTALYA HÝZMETLERÝ</span>
<a href="antalya-klima-servisi.asp"><div><i class="fas fa-chevron-right"></i> Antalya Klima Servisi</div></a>
<a href="dosemealti-klima-servisi.asp"><div><i class="fas fa-chevron-right"></i> Dosemealti Klima Servisi</div></a>
<a href="antalya-klima-montaji.asp"><div><i class="fas fa-chevron-right"></i> Antalya Klima Montaji</div></a>
<a href="antalya-klima-bakimi.asp"><div><i class="fas fa-chevron-right"></i> Antalya Klima Bakimi</div></a>
<a href="antalya-mitsubishi-electric-klima.asp"><div><i class="fas fa-chevron-right"></i> Mitsubishi Electric Klima</div></a>
<a href="antalya-vrf-sistemleri.asp"><div><i class="fas fa-chevron-right"></i> Antalya VRF Sistemleri</div></a>

<span class="solMenuBas" style="margin-top:18px; display:block;" >KLÝMA REHBERÝ</span>
<a href="klima-rehberi.asp"><div><i class="fas fa-chevron-right"></i> Klima Rehberi Blog</div></a>