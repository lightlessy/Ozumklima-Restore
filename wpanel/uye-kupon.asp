<html>
<head>    
<title>Müşteri Hesabı Düzenle</title>
  <meta http-equiv="Content-Type" content="text/html; charset=windows-1254">        
<link rel=stylesheet type="text/css" href="font.css">
</head>   

 
<!--#INCLUDE file="dbase.asp"-->   

                
</HEAD>               

 
 <% if request("add")<>"" then  

 
set Rs = Server.CreateObJect("ADODB.RecordSet")
Sorgula = "Select * From indir"
Rs.open Sorgula,baglanti,1,3    

Rs.AddNew            
Rs("kodu")=   Request.Form("kodu") 
Rs("bittar")=   Request.Form("bittar") 
Rs("oran")=   Request.Form("oran") 
Rs("uid")=   Request.Form("uid") 
Rs.Update      

rs.close
set rs = nothing    
Response.Redirect Request.ServerVariables("HTTP_REFERER")
  end if           
  
  
   if request("delete")<>"" then 
   
  Set Rs = Server.CreateObject("Adodb.Recordset")      
rSQL = "Select * from indir where  id = "& request("id") &" "
Rs.Open rSQL, baglanti, 1, 3 
     
Rs.Delete 
Rs.Update     

Response.Redirect("uye-kupon.asp?id="& request("uid") &"")

end if      
  %>

       
<body>  

<%
Set rs = Server.CreateObject("ADODB.Recordset")
sql = "SELECT * FROM uyeler WHERE uid = "& request("id")
rs.open sql, baglanti,1,3    
%>     

               <BR>


<table  bgcolor=#F5F5F5 bordercolor=#EAEAEA cellpadding=5 cellspacing=5 align=center class=font width=400 border=1> 

<form method="POST" action="uye-kupon.asp?id=request("id")" name="myForm" id="searchform">       

<tr><td><b>Ad Soyad</b></td><td><%=rs("ad")%>&nbsp;<%=rs("soyad")%></td></tr>
<tr><td><b>İndirim Kodu</b></td><td><input type="text" size="14" name="kodu" maxlength="8" >  </td>    </tr>
<tr><td><b>Son Geçerlilik Tarihi</b></td><td><input type="date" size="14" name="bittar"  >  </td>    </tr>  
<tr><td><b>İndirim Oranı</b></td><td>%<input type="text" size="5" name="oran" maxlength="2" >  </td>    </tr>


<input type="hidden" name="uid"   value="<%=rs("uid")%>">

<input type="hidden" name="add"    value="1">
 
<tr><td colspan=4 align=center>
<input type="submit" name="update" value="Tanımla"></td>

</form></tr>
</table>      

<BR> <BR>

 

<table class=font align=center bgcolor=#F5F5F5 bordercolor=#EAEAEA cellpadding=3 cellspacing=3 width=400 border=1>   
<tr><td colspan=6 align=center> <B>Üyenin Tüm Kodları</B></td></tr>     

<tr><td><b>İndirim Kodu</b></td><td><b>Son Geçerlilik Tarihi</b></td><td><b>İndirim Oranı</b></td> <td align=center><B>Sil</B></td></tr>

<%
Set object = Server.CreateObject("ADODB.Recordset")
sql = "SELECT * FROM indir WHERE uid = "& request("id")
object.open sql, baglanti,1,3    

Do while not object.Eof                                                                                                                                                                                                                  
%>  


<tr>
<td><%=object("kodu")%> </td>  
<td><%=object("bittar")%>  </td>    
<td>% <%=object("oran")%>  </td>
<td><a href="uye-kupon.asp?id=<%=object("id")%>&uid=<%=object("uid")%>&delete=1"><img src="images/del.png" width=20></a>  </td>     </tr> 
  

<%  
object.MoveNExt
Loop
%>
  </table>       


           
      </body>
      </html>