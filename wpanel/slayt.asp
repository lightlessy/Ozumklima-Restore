
  <!--#include file="menu.asp"-->


<%
Set UploadProgress = Server.CreateObject("Persits.UploadProgress")
PID = "PID=" & UploadProgress.CreateProgressID()
barref = "framebar.asp?to=10&" & PID
%>

<SCRIPT LANGUAGE="JavaScript">
function ShowProgress()
{
  strAppVersion = navigator.appVersion;
  if (document.MyForm.resim.value != "" || document.MyForm.FILE2.value != "" || document.MyForm.FILE3.value != "")
  {
    if (strAppVersion.indexOf('MSIE') != -1 && strAppVersion.substr(strAppVersion.indexOf('MSIE')+5,1) > 4)
    {
      winstyle = "dialogWidth=385px; dialogHeight:140px; center:yes";
      window.showModelessDialog('<% = barref %>&b=IE',null,winstyle);
    }
    else
    {
      window.open('<% = barref %>&b=NN','','width=375,height=115', true);
    }
  }
  return true;
}
</SCRIPT>


  <body>

  <div class="baslik">Slayt Ekle </div>


<BR><BR>


	<form method="post" NAME="MyForm" action="upload_flas.asp?<% = PID %>" ENCTYPE="multipart/form-data" OnSubmit="return ShowProgress();">



 <table bgcolor=#FDFDFD cellpadding=5 cellspacing=5 align=center width=700 class=font align=center>

<tr><td><b>Görsel Seç </b></td><td><input type="file" name="resim" size="51"></td>	</tr>
 <tr><td><b>Gideceği Bağlantı </b></td><td><input type="text" size="80" name="adres"></td></tr>

  <tr><td colspan=2 align=center><input type="submit" value="YÜKLE" class="admin_button"></td>
</td></tr>

</table>
      </form>



