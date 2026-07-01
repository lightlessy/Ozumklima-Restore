   <span class="solMenuBas" >KURUMSAL</span>

   <%  

set object = Server.CreateObJect("ADODB.RecordSet")
Sorgula = "Select * From sayfalar where ana=1 order by sira asc"
object.open Sorgula,baglanti,1,3  

Do while not object.Eof                                                                                                                                                                                                                  
%>   

 <a href="kurumsal.asp?id=<%=object("id") %>"  ><div><i class="fas fa-chevron-right"></i> <%=object("isim") %></div></a>   
 
<%  
object.MoveNExt
Loop
%>