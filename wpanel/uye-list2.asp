<html>
 <head>
  <title>PANEL </title> 
  <meta name="viewport" content="width=device-width, initial-scale=1">         

  <meta http-equiv="Content-Type" content="text/html; charset=windows-1254">        
<link rel=stylesheet type="text/css" href="font.css"> 

           

</head>     

<!--#INCLUDE file="dbase.asp"-->      

  <script src="../js/jquery-2.1.1.js"></script>   
 

</HEAD>   


 <%if request.form("delete")<>"" then  %>

<%

Set Rs = Server.CreateObject("Adodb.Recordset")      
rSQL = "Select * from uyeler where uid = "& Request.Form("uid")
Rs.Open rSQL, baglanti, 1, 3 
Rs.Delete 
Rs.Update 

rs.close
set rs = nothing 
Response.Redirect Request.ServerVariables("HTTP_REFERER")
%>        
<% end if %>       


<body>         
   

<div id="postswrapper">
<div class="postswrapper">

<table  cellpadding=5 cellspacing=4 align=center class=font border=0 width=720> 
 
         

          
<%  
 id =request.form("id")
 
set uyes = Server.CreateObJect("ADODB.RecordSet")
Sorgula = "Select top 20 * From uyeler where uid < "& id &"  order by uid desc "
uyes.open Sorgula,baglanti,1,3    
        
i=1

Do while not uyes.Eof   

if i mod 2 then 
bg="#FFFFFF" 
else
bg="#F2F2F2" 
end if  %>    
               

<ul style="list-style:none; ">  
  
<li id="<%=uyes("uid")%>"> 

<tr height=35 bgcolor=<%=bg%>>
<td width=20><B><%=uyes("uid")%></b></td>   
<td width=150><B><%=uyes("ad")%>&nbsp;<%=uyes("soyad")%></td>    


            

<div id="inline<%=uyes("uid")%>" style="background:#fff; padding:15px; font-family:Open Sans; font-size:14px; color:#373737; width:500px; height:400px  " class="lity-hide">

<b><%=uyes("ad")%>&nbsp;<%=uyes("soyad")%> </b>   <BR> <BR>     

<%  
set mods1 = Server.CreateObJect("ADODB.RecordSet")
Sorgula = "Select tarih, kod, uid, tfiyat, onay From arsiparis where uid= "& uyes("uid") & " and onay=1 group by  tarih, kod,uid, tfiyat, onay order by tarih desc"
mods1.open Sorgula,baglanti,1,3         

 If mods1.EOF And mods1.BOF Then %>     

 <BR><BR> <B> SİPARİŞ YOK </B>         

<% else %>

<% Do while not mods1.Eof  %>        

 <a href="javascript:void(window.open('ar_order_det1.asp?id=<%=mods1("kod")%>','','width=900,height=600,left=400,top=0,resizable=no,menubar=no,centerscreen=yes,status=no,scrollbars=1'))" style="color:#00557D">
<B><%=mods1("kod")%></B>  - <%=mods1("tarih")%> - <%=mods1("tfiyat")%> TL  <B> » Detay   </b></a>  
<hr style="height: 1px;background-color: #B9B9B9; border:none;">       

 <%  
mods1.MoveNExt
Loop
%>     

<% end if %>
</div>

<td align=center width=70>  <a href="#inline<%=uyes("uid")%>" data-lity><font color=#800000><B>Siparişler</B> </font></a></td> 



<div id="indline<%=uyes("uid")%>" style="background:#fff; padding:15px; font-family:Open Sans; font-size:14px; color:#373737; width:500px; height:650px  " class="lity-hide">
<IFRAME frameBorder=0 scrolling=no width=500 height=500 src="uye-kupon.asp?id=<%=uyes("uid")%>" ></IFRAME> 
</div>
<td align=center width=100> <a href="#indline<%=uyes("uid")%>" data-lity>İndirim Kuponu  </a></td>


<div id="inline2<%=uyes("uid")%>" style="background:#fff; padding:15px; font-family:Open Sans; font-size:14px; color:#373737; width:500px; height:650px  " class="lity-hide">
<IFRAME frameBorder=0 scrolling=no width=500 height=500 src="uye-edit.asp?id=<%=uyes("uid")%>" ></IFRAME> 
</div>
<td align=center width=70> <a href="#inline2<%=uyes("uid")%>" data-lity><img src="images/edit.png" width=20></a></td>


<form action="uye-list.asp" method="post" >
<td align=center width=50>
<input type="hidden" name="uid" value="<%=uyes("uid")%>">        
<input type="hidden" name="delete" value="1">     
<input type="image" src="images/del.png"  onclick="return confirm('Müşteri Hesabını Silmek İstiyor Musunuz?')">
</td>

</form>
</tr>   

</li>    
 </ul>


<%        
  i=i + 1
uyes.MoveNExt
Loop
%>     
   


				</div>       
				</div>

</table>          

    
 
 
 
           
      </body>
      </html>