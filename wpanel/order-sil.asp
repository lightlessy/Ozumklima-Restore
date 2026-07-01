

   
  <!--#include file="menu.asp"-->  


  
<%


Set Rs = Server.CreateObject("Adodb.Recordset")
rSQL = "Select * from siparis  where kod = '"& Request("sil") &"'"
Rs.Open rSQL, baglanti, 1, 3 


Do while not Rs.Eof                                                                                                                                                                                                                     


Rs.Delete 
Rs.Update 


Rs.MoveNExt
Loop

rs.close
set rs = nothing 
Response.Redirect Request.ServerVariables("HTTP_REFERER")         

    %>   


