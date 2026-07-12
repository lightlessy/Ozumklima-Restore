<%@ Language="VBScript" CodePage="65001" %>
<%
Response.Buffer = True
Response.CodePage = 65001
Response.Charset = "utf-8"
%>

<html>
<head>
<title>Müşteri Hesapları</title>


<!--#include file="menu.asp"-->


                  <link rel="stylesheet" href="../css/lity.css">
                  <script src="../js/lity.js"></script>

</HEAD>






<script type="text/javascript">
$(window).scroll(function()
{
    if((window.innerHeight + window.scrollY  ) >= document.body.offsetHeight) {



        $('div#loadmoreajaxloader').show();

        var id = $(".postswrapper ul li:last").attr("id");

        $.ajax({
        type: "post",
        url: "uye-list2.asp",
        data: {"id":id},
        success: function(html)
        {
            if(html)
            {
                $("#postswrapper").append(html);
                $('div#loadmoreajaxloader').hide();
            }else
            {
                $('div#loadmoreajaxloader').html('<center>No more posts to show.</center>');
            }
        }
        });
    }
});
</script>

<div class="yuks"></div>

	<style>


		#wrapper{
			width:600px;
			margin:auto;
		}
		.spacer{
			clear:both;
			height:5px;
		}
		.txtarea{
			font-size:18px;
			height:50px;
			width:100%;
		}
		#postswrapper{
			 list-style: none;

		}

	</style>






  <div class="baslik">Müşteri Hesapları</div>

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
<BR>




<div id="postswrapper">
<div class="postswrapper">

<table  cellpadding=5 cellspacing=4 align=center class=font border=0 width=580>
 <tr bgcolor=#F3F0E4 height=28>
 <td>ID </td>
  <td width=200>Ad Soyad </td>
<td align=center width=70><B>Siparişler</B></td>
<td align=center width=100><B>İndirim</B></td>
<td align=center width=70><B>Düzenle</B></td>
<td align=center width=50><font color=#CA0000><B>Sil</B> </td>
    </tr>


<%

Set Rs = Server.CreateObject("Adodb.Recordset")
rSQL = "Select * from uyeler  "
Rs.Open rSQL, baglanti, 1, 3

Do while not Rs.Eof
 if   Rs.fields("kullanici")&"" = "" then
Rs.Delete
Rs.Update
end if

Rs.MoveNExt
Loop


%>


<%
set uyes = Server.CreateObJect("ADODB.RecordSet")
Sorgula = "Select top 30 * From uyeler   order by uid desc "
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
<td width=10><B><%=uyes("uid")%></b></td>
<td width=140><B><%=uyes("ad")%>&nbsp;<%=uyes("soyad")%></td>




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

<td align=center>  <a href="#inline<%=uyes("uid")%>" data-lity><font color=#800000><B>Siparişler</B> </font></a></td>



<div id="indline<%=uyes("uid")%>" style="background:#fff; padding:15px; font-family:Open Sans; font-size:14px; color:#373737; width:500px; height:650px  " class="lity-hide">
<IFRAME frameBorder=0 scrolling=no width=500 height=500 src="uye-kupon.asp?id=<%=uyes("uid")%>" ></IFRAME>
</div>
<td align=center> <a href="#indline<%=uyes("uid")%>" data-lity>İndirim Kuponu  </a></td>


<div id="inline2<%=uyes("uid")%>" style="background:#fff; padding:15px; font-family:Open Sans; font-size:14px; color:#373737; width:500px; height:650px  " class="lity-hide">
<IFRAME frameBorder=0 scrolling=no width=500 height=500 src="uye-edit.asp?id=<%=uyes("uid")%>" ></IFRAME>
</div>
<td align=center> <a href="#inline2<%=uyes("uid")%>" data-lity><img src="images/edit.png" width=20></a></td>


<form action="uye-list.asp" method="post" >
<td align=center>
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

<tr><td colspan=6>
	<div id="loadmoreajaxloader" style="display:none;"><center><img src="images/lightbox-ico-loading.gif" /><BR><BR><BR></center></div>
</td></tr>
				</div>
				</div>




</table>







      </body>
      </html>