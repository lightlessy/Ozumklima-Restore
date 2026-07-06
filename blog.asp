<%
Dim idParam, id, Sorgu, icerisi, icerikVar

idParam = Trim(Request.QueryString("id") & "")

If IsNumeric(idParam) Then
    id = CLng(idParam)
Else
    id = 0
End If

icerikVar = False
%>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html>
<head>  

<META HTTP-EQUIV="Content-Type" content="text/html; charset=windows-1254"> 
<meta name="viewport" content="width=device-width, initial-scale=1">

<title>Blog | Antalya Klima - &Ouml;z&uuml;m Klima Mitsubishi Klima Antalya</title>    

<meta name="author" content="Mega Tasar&#305;m"/>   

<link rel="icon" type="image/png" href="fav.png" />   

<link href="css/font.css" rel="stylesheet" type="text/css"/>   

<script src="js/jquery-2.1.1.js"></script>  
<script src="js/jquery-ui.min.js"></script>               

<link href="css/fontawesome-all.css" rel="stylesheet">
<link href="css/lity.css" rel="stylesheet">

<script src="js/lity.min.js"></script>

</head>

<body> 

<!--#include file="ust.asp"-->

<%
Set icerisi = Server.CreateObject("ADODB.RecordSet")

If id > 0 Then
    Sorgu = "Select * From sayfalar where id=" & id & " and ana=3"
Else
    Sorgu = "Select * From sayfalar where ana=3 order by sira asc"
End If

icerisi.Open Sorgu, baglanti, 1, 1

If Not icerisi.EOF Then
    icerikVar = True
End If
%>

<Div class="tepeRestas" style="background-image:url('images/arka-kurumsal.jpg')"></Div>

<DIV class="MenuVeicerik">  

    <Div class="solMenu">
        <!--#include file="sol-blog.asp"-->
    </Div>
      
    <Div class="icerik2" style="white-space: pre-line;">
        <% If icerikVar = True Then %>

            <span><B><%=icerisi("isim") %></B></span>  
            <%=icerisi("icerik") %>

        <% Else %>

            <span><B>&#304;&#231;erik bulunamad&#305;.</B></span>
            Blog kategorisinde i&#231;erik bulunamad&#305;.

        <% End If %>
    </Div>     

    <div style="clear:both"></div>

</DIV>             

<!--#include file="alt.asp"-->      

</body>
</html>

<%
If IsObject(icerisi) Then
    If icerisi.State = 1 Then icerisi.Close
    Set icerisi = Nothing
End If
%>