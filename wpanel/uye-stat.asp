<html>
<head>                    
  
 <script language="javascript" src="cal2.js"></script>
<script language="javascript" src="cal_conf2.js"></script>

<title>Müþteri Raporlarý</title>     

<% sira= request.querystring("sira")
tp= request.querystring("tp")
if sira ="" then
sira ="toplam"
end if 
if tp ="" then
tp ="desc"
end if  
t1=Request.form("t1")  
t2=Request.form("t2")  %>

<!--#include file="menu.asp"-->      

</HEAD>               


<body>  
              <BR>
  
<center><BR><font class=fontbig><b>Müþteri Raporlarý</b></font>   <BR><BR> 

 
   
<form method="POST" action="uye-stat.asp" name="sampleform">       
   
               
<input name="t1"   size="12" value="<%=t1%>"><a href="javascript:showCal('Calendar1')"> <img src="cal.gif" width="16" height="16" border="0" alt="Select Date"></a>
--
<input name="t2"   size="12" value="<%=t2%>"><a href="javascript:showCal('Calendar2')"> <img src="cal.gif" width="16" height="16" border="0" alt="Selecet Date"></a>
<input type="submit" name="1" value=">>" />
   
  </form>              

<table  cellpadding=4 cellspacing=5 align=center class=font border=0 width=920> 
 <tr bgcolor=#F2F2F2>    
 <td width=170><b>Ad Soyad </td>
 <td width=170><b>Telefon </td>
 <td width=170><b>Adres </td>
<td align=center><b>Toplam Sipariþ </td>   
                                                       
    </tr>
             
<% if request.form("1") <>"" then %>

<%        

set foods = Server.CreateObJect("ADODB.RecordSet")
Sorgula = "Select uid From arsiparis where gun between CDate('"&t1&"')and CDate('"&t2&"') group by uid"
foods.open Sorgula,baglanti,1,3      

i=1

Do while not foods.Eof   

if i mod 2 then 
bg="#FFFFFF" 
else
bg="#F2F2F2" 
end if          
    
Set rs1 = baglanti.Execute("Select SUM (tfiyat) AS toplam FROM arsiparis where uid="& foods("uid") &" and gun between  CDate('"&t1&"')and CDate('"&t2&"') ; ")        

set uye = Server.CreateObJect("ADODB.RecordSet")
Sorgula = "Select * From arsiparis where uid = "& foods("uid") &" and  gun between  CDate('"&t1&"')and CDate('"&t2&"') "
uye.open Sorgula,baglanti,1,3      


set uyes = Server.CreateObJect("ADODB.RecordSet")
Sorgula = "Select * From uyeler where uid = "& foods("uid") &"  "
uyes.open Sorgula,baglanti,1,3   

Do while not uyes.Eof   
%>            
  
<tr height=35 bgcolor=<%=bg%>>
<td width=100><B><%=uyes("ad")%>&nbsp;<%=uyes("soyad")%>  </td>   
<td width=120><B>Gsm: <%=uyes("cep")%><BR></b> Tel.: <%=uyes("tel")%> </td>  
<td><B><%=uyes("adres")%></td>  
<td align=center>   <B><%=rs1("toplam")%>  TL  </B></td>  


</tr>   


<%        

uyes.MoveNExt
Loop
%>      
         

<%        
  i=i + 1
 foods.MoveNExt
Loop     
%>

</table>          
                    

<% end if %>

           
      </body>
      </html>