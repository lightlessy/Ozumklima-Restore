<!DOCTYPE html> 
<html lang="en"> 
    <head>
<meta charset="utf-8">
<meta http-eqiv="Content-Type" content="windows-1254">      

        <title>Resim Yükle</title>
      <meta name="apple-mobile-web-app-capable" content="yes" />
        <meta name="viewport" content="width=device-width, initial-scale=1.0, maximum-scale=1.0" />
     
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

    </head>
    <body>    
    
    
    
    <!--#INCLUDE file="dbase.asp"-->  
    
      <!--#include file="menu.asp"-->  
      

    
  
             
 
  
  <%
  id = request.querystring("id")      
g = request.querystring("g")    
n = request.querystring("n")  
tip = request.querystring("tip")
isim = request.querystring("isim")  

%> 
     
   <BR>
   
<div style="  text-align:center; margin-top:20px">  
    

	<FORM NAME="MyForm" METHOD="POST" ENCTYPE="multipart/form-data"  ACTION="imageupdate2.asp?<% = PID %>"	OnSubmit="return ShowProgress();"    style="margin-top:15px; display: inline-block; padding: 15px; text-align: center; background-color:#314F51; color:#FFFFFF;  border-radius: 6px;"> 
	     Fotoğraf:	<INPUT TYPE="FILE" SIZE="40" NAME="foto" required><BR>  <BR>   
        
<input  type="hidden" name="isim" value="<%=isim%>"> 	     
<input  type="hidden" name="id" value="<%=id%>">   
<input  type="hidden" name="tip" value="<%=tip%>">  
<input  type="hidden" name="n" value="<%=n%>">   
<input  type="hidden" name="g" value="<%=g%>">    
<input  type="hidden" name="git" value="2"> 
	<INPUT TYPE=SUBMIT VALUE="YÜKLE">
	</FORM>                   
    
  </div>  
  
               
                 
     
     

           

 
       
                                              
    </body>
</html> 
