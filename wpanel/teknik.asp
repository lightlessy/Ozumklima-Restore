    <!--#INCLUDE file="dbase.asp"-->  
    
      <!--#include file="menu.asp"-->  

<HEAD>



<%
Set UploadProgress = Server.CreateObject("Persits.UploadProgress")
PID = "PID=" & UploadProgress.CreateProgressID()
barref = "framebar.asp?to=10&" & PID
%>

<SCRIPT LANGUAGE="JavaScript">
function ShowProgress()
{
  strAppVersion = navigator.appVersion;
  if (document.MyForm.foto2.value != "" || document.MyForm.FILE2.value != "" || document.MyForm.FILE3.value != "")
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



                 <center><BR><BR>        <BR>
                 
                 <div style="width:350px; background-color:#FFFFFF; border:1px solid #E8E8E8; padding:15px; text-aligfn:center;">                        
			<FORM NAME="MyForm" METHOD="POST" ENCTYPE="multipart/form-data" 
		ACTION="teknik2.asp?<% = PID %>"
		OnSubmit="return ShowProgress();"> 
		<INPUT TYPE="FILE" SIZE="40" NAME="dokuman" style="background-color:#E8E8E8"><BR>   
		 
<input type="hidden" name="id" value="<%=request.querystring("id")%>">
<input type="hidden" name="g" value="<%=request.querystring("g")%>">
<input type="hidden" name="tip" value="<%=request.querystring("tip")%>">
		    <BR><BR>	
		 	
	<INPUT TYPE=SUBMIT VALUE="YÜKLE">
	</FORM> 
	
	</div>

         <BR><BR><BR>

      </body>
      </html>