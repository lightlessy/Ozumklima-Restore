<!--#INCLUDE file="dbase.asp"-->   


<%       

toplam = Cint(Request.Form("toplam"))     



for i = 1 to toplam 
if Request.Form("Veri"&i) <> "" then 
Set Rs = Server.CreateObject("Adodb.Recordset")
rSQL = "Select * from user  where id = "& Request.Form("Veri"&i)
Rs.Open rSQL, baglanti, 1, 3 

Rs.Delete ' Sil
Rs.Update ' Güncelle

rs.close
set rs = nothing 
end if
Next

baglanti.close
set baglanti = nothing

Response.Redirect Request.ServerVariables("HTTP_REFERER") 

%>    

