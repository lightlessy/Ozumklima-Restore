   
<!--#INCLUDE file="menu.asp"-->

<body>   

<%


if Request.Querystring("psil2")=1 then 


Set Rs = Server.CreateObject("Adodb.Recordset")
rSQL = "Select * from products  where AffiliateID = "& Request("id")
Rs.Open rSQL, baglanti, 1, 3 

Rs.Delete 
Rs.Update 

rs.close
set rs = nothing 
Response.Redirect Request.ServerVariables("HTTP_REFERER")         

end if      %>   


<div class="baslik">Ürün Bul</div>

<BR><BR>    

<%                                 
If request.form("ara") <> "" then 


	SQLstring = "SELECT * FROM products WHERE "& request.form("field") &" LIKE '%"& request.form("search") &"%' order by AffiliateID desc"
	set RsSearch = Server.CreateObject("ADODB.Recordset")
RsSearch.Open SQLstring, baglanti, 1, 3 
End if      
%>


  

<form name="form1" method="post" action="urun_bul.asp">
  <table width="610" border=1 bordercolor=#EEEEEE cellspacing="5" cellpadding="5" class="font" align=center bgcolor=#FFFFFF>
    <tr> 
      <td align=center>                                                                                
        <select name="field">
<option value="isim">Ürün Adı</option> 
 </select>
        <input type="text" name="search">
                <input type="hidden" name="ara" value="1">
</td>
    </tr>
    <tr bgcolor=#EEEEEE> 
      <td align=center> <input type="submit" name="Submit" value="A R A">

    </td></tr>
  </table>   
 </form>
 
 
<%
If request.form("ara") <> "" then 
%>   <font face=arial size=2 color=#444444><B>              
<% if rssearch.eof then
response.write "KAYIT BULUNAMADI"
else
%>




<table border="1" bordercolor="#E6E6E6" bgcolor=#FFFFFF cellpadding="3" cellspacing="3" width=700 class=font align=center>         
<tr><td colspan=3 align=center>
<B>ARAMA SONUÇLARI</b>   
</td></tr>
<tr bgcolor=#F3F0E4 height=30>    
   <td><b>Ürün Adı </b></td>    

<td align=center><font color=#0080C0><b>Düzenle</b></td><td align=center><font color=#FF0000><b>SİL</b></td></tr>      


<%
  Do while not rssearch.Eof   
%>
             


<tr class="menuoff" onmouseover="className='menuon';" onmouseout="className='menuoff';">
<td><%=rssearch("isim")%></td>
<td align=center><a href="edits.asp?id=<%=rssearch("AffiliateID")%>&g=<%=rssearch("anagrup")%>&tip=<%=rssearch("tip")%>"><img src="images/edit.png" width=20></a><td align=center> <a href="?psil2=1&id=<%=rssearch("AffiliateID")%>" onclick="return confirm('Kayıt silinecek. Onaylıyor Musunuz?')"><img src="images/del.png" width=20></a>
</td>      
</tr>


<%
rssearch.MoveNExt
Loop
%>        

   </tr></table>      
   
   <% end if %> 

<% end if %> 

      <BR><BR>


 