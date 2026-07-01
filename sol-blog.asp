   <span class="solMenuBas" >KONFORLOG</span>
 
 <% 
Function cevir(data)   
if not data="" then
data = Replace (data ,"ý","i",1,-1,1) 
data = Replace (data ,"ð","g",1,-1,1) 
data = Replace (data ,"ü","u",1,-1,1) 
data = Replace (data ,"þ","s",1,-1,1) 
data = Replace (data ,"ö","o",1,-1,1) 
data = Replace (data ,"ç","c",1,-1,1) 
data = Replace (data ,"I","i",1,-1,1) 
data = Replace (data ,"Ð","g",1,-1,1) 
data = Replace (data ,"Ü","u",1,-1,1) 
data = Replace (data ,"Þ","s",1,-1,1) 
data = Replace (data ,"Ý","i",1,-1,1) 
data = Replace (data ,"Ç","c",1,-1,1) 
data = Replace (data ,"Ö","o",1,-1,1) 
data = Replace (data ," ","-",1,-1,1)   
end if
cevir=data 
End Function 
%> 

   <%  

set object = Server.CreateObJect("ADODB.RecordSet")
Sorgula = "Select * From sayfalar where ana=3 order by sira asc"
object.open Sorgula,baglanti,1,3  

Do while not object.Eof                                                                                                                                                                                                                  
%>   

 <a href="konforlog.asp?k=<%=cevir(object("isim")) %>&id=<%=object("id") %>"  ><div><i class="fas fa-chevron-right"></i> <%=object("isim") %></div></a>   
 
<%  
object.MoveNExt
Loop
%>