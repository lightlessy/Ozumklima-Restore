

<html>
<head>   
<meta charset="utf-8">
<script language=javascript>

function CloseWindow()

{

window.open('','_self','');

window.close();

}

</script>            

   
<!--#INCLUDE file="menu.asp"-->        

                           

<BR><BR><center><font face=verdana size=2><%

set id = Request.QueryString("id")
sql = "delete from galery where id="
sql = sql & id
set db=baglanti.Execute(sql)    

response.redirect("galeri_edit.asp?gal_kat="&request("gal_kat")&"&altkat="&request("altkat")&"")
%>  
   



</body>
</html>

