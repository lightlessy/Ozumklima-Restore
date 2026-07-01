<html>
<head>
<title>Ýptal Sipariþler</title>

  <!--#include file="menu.asp"-->                     
               
  
</HEAD>   


<body>  

  <div class="baslik">Ýptal Sipariþler </div>   



<TABLE border=0 class=font cellspacing=3 cellpadding=4 align=center><TR>
<TD valign=top>   
             
<%        



set mods1 = Server.CreateObJect("ADODB.RecordSet")
Sorgula = "Select gun,onay, kod From arsiparis  where  onay=1 group by gun,onay,kod  order by gun desc"
mods1.open Sorgula,baglanti,1,3      
   
%>   
              

<TD valign=top>        
                              

<% Do while not mods1.Eof  %>        


<BR>
<table bgcolor=#f3f7f8 cellpadding=2 cellspacing=2 align=center class=font border=0 width=500>   
     
 <tr bgcolor=#FFFFFF>
 <td colspan=2 class=onay>  
 <a href="iptal_order_det.asp?gun=<%=mods1("gun")%>"><font color=#BF0000>  <B><%=mods1("gun")%></B></font><font color=#4E4E4E> &nbsp; Tarihli Sipariþler   » </a>    </font>              
</td></tr>              


</table> 

    <%  
mods1.MoveNExt
Loop
%>     

</TD></TR></TABLE>   






      </body>
      </html>