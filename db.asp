

<%
set baglanti=server.createobject("adodb.connection")
baglanti.open "driver={microsoft access driver (*.mdb)}; dbq=" &server.mappath("db/TRvd678577.mdb")   


Set ayars = Server.CreateObject("ADODB.Recordset")
sqly = "SELECT * FROM ayarlar"
ayars.open sqly, baglanti,1,3    


set car = Server.CreateObJect("ADODB.RecordSet")
Sorgula = "Select * From cargo "
car.open Sorgula,baglanti,1,3



     
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
data = Replace (data ,"%","",1,-1,1)   
end if
cevir=Lcase(data)
End Function 




Function cevir1(data)   
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
data = Replace (data ,"%","",1,-1,1) 
end if
cevir1=Ucase(data)
End Function 
                     
                        
%>   
           
