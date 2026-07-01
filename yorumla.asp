
<!--#include file='db.asp'-->    
<%

     

            
   
 
      SET rs = Server.CreateObject("ADODB.Recordset")        
      sql = "select * from yorumlar"     
      rs.open sql,baglanti,1,3 

Rs.AddNew            
Rs("yorum")=   Request.Form("yorum") 
Rs("isim")=   Request.Form("isim")       
Rs("onay")=  0
Rs.Update       
	  rs.close       
	  

 response.redirect("ok.asp")     
                                                                                                                               


%>