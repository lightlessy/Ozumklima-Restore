<!--#INCLUDE file="menu.asp"-->      

<body>



<%   

sil=Request.querystring("sil") 
id=Request.querystring("id") 
                                                  
  

   Set Upload = Server.CreateObject("Persits.Upload")             
   Directory = ""&Server.MapPath("../urunler")&"/"  
         
   Upload.DeleteFile Directory & sil   



Set Rs = Server.CreateObject("Adodb.Recordset")
rSQL = "Select * from  products where AffiliateID = "& id &" "
Rs.Open rSQL,baglanti,1,3

Rs("foto1") = Request("foto1")       
Rs("foto2") = Request("foto2")       
Rs("foto3") = Request("foto3")       
Rs("foto4") = Request("foto4")       
Rs("foto5") = Request("foto5")       

Rs.Update 
rs.close
set rs = nothing   
                                  
Response.Redirect Request.ServerVariables("HTTP_REFERER")

%>



 
</body>
</html>
