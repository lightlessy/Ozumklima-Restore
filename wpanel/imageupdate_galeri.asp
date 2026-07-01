
<HEAD>

<title>Foto Güncelle </title>

<%
Set UploadProgress = Server.CreateObject("Persits.UploadProgress")
PID = "PID=" & UploadProgress.CreateProgressID()
barref = "framebar.asp?to=10&" & PID
%>

<SCRIPT LANGUAGE="JavaScript">
function ShowProgress()
{
  strAppVersion = navigator.appVersion;
  if (document.MyForm.foto.value != "" || document.MyForm.FILE2.value != "" || document.MyForm.FILE3.value != "")
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

</HEAD>

<!--#INCLUDE file="menu.asp"-->    

    <div class="baslik">Fotoðraf Güncelle</div>     

              		<% Session("id") = request.querystring("id")  
              		Session("gal_kat") = request.querystring("gal_kat")           %>       

                       <BR><BR>                                                                  <CENTER>        
                       
                        <table bgcolor=#F7F7F7 cellpadding=5 cellspacing=5 align=center class=fontsm> 

            <tr>	<td colspan=2>
                                                 
			<FORM NAME="MyForm" METHOD="POST" ENCTYPE="multipart/form-data" 
		ACTION="updateimage_galeri.asp?<% = PID %>&gal_kat=<%=request("gal_kat")%>"
		OnSubmit="return ShowProgress();"> 
	     Fotoðraf:	<INPUT TYPE="FILE" SIZE="40" NAME="foto">
	     
	     </td></tr>
  <tr><td colspan=2 align=center> 
	<INPUT TYPE=SUBMIT VALUE="YÜKLE">
	</FORM>               
	</td></tr></table>    
</CENTER>

               
<BR><BR><BR>

      </body>
      </html>