<% @ Language=VBScript %>
<% Option Explicit %> 



<!--#include file="RTE_configuration/RTE_setup.asp" -->
<!--#include file="functions/RTE_functions_common.asp" -->
<!--#include file="language_files/RTE_language_file_inc.asp" -->
<%
'****************************************************************************************
'**  Copyright Notice    
'**
'**  Web Wiz Rich Text Editor(TM)
'**  http://www.richtexteditor.org
'**                                                              
'**  Copyright (C)2001-2012 Web Wiz Ltd. All Rights Reserved.     
'**  
'**  THIS SOFTWARE IS PROVIDED BY THE COPYRIGHT HOLDERS UNDER LICENSE FROM WEB WIZ LTD.
'**  
'**  IF YOU DO NOT AGREE TO THE LICENSE AGREEMENT THEN WEB WIZ LTD. IS UNWILLING TO LICENSE 
'**  THE SOFTWARE TO YOU, AND YOU SHOULD DESTROY ALL COPIES YOU HOLD OF 'WEB WIZ' SOFTWARE
'**  AND DERIVATIVE WORKS IMMEDIATELY.
'**  
'**  If you have not received a copy of the license with this work then a copy of the latest
'**  license contract can be found at:-
'**
'**  http://www.webwiz.co.uk/license
'**
'**  For more information about this software and for licensing information please contact
'**  'Web Wiz' at the address and website below:-
'**
'**  Web Wiz Ltd, Unit 10E, Dawkins Road Industrial Estate, Poole, Dorset, BH15 4JD, England
'**  http://www.webwiz.co.uk
'**
'**  Removal or modification of this copyright notice will violate the license contract.
'**
'****************************************************************************************



'*************************** SOFTWARE AND CODE MODIFICATIONS **************************** 
'**
'** MODIFICATION OF THE FREE EDITIONS OF THIS SOFTWARE IS A VIOLATION OF THE LICENSE  
'** AGREEMENT AND IS STRICTLY PROHIBITED
'**
'** If you wish to modify any part of this software a license must be purchased
'**
'****************************************************************************************




'Set the response buffer to true
Response.Buffer = True



'Dimension veriables
Dim strOpenFilePath

'Get the open file path
strOpenFilePath = strOpenFileFolderPath

'Change \ for /
strOpenFilePath = Replace(strOpenFilePath, "\", "/'", 1, -1, 1)
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta charset="utf-8">
<title>Open File</title>
<!--#include file="RTE_configuration/browser_page_encoding_inc.asp" -->

<%
'***** START WARNING - REMOVAL OR MODIFICATION OF THIS CODE WILL VIOLATE THE LICENSE AGREEMENT ******
Response.Write(vbCrLf & "<!--//" & _
vbCrLf & "/* *******************************************************" & _
vbCrLf & "Software: Web Wiz Rich Text Editor(TM) ver. " & strRTEversion & "" & _
vbCrLf & "Info: http://www.richtexteditor.org" & _
vbCrLf & "Copyright: (C)2001-2012 Web Wiz Ltd. All rights reserved" & _
vbCrLf & "******************************************************* */" & _
vbCrLf & "//-->")
'***** END WARNING - REMOVAL OR MODIFICATION OF THIS CODE WILL VIOLATE THE LICENSE AGREEMENT ******
%>

<script language="JavaScript">
//function to open the file
function getFile(){
	var file = document.getElementById('loader').contentWindow.document.body.innerHTML.toString();
	if (file != ''){
		window.opener.document.getElementById('WebWizRTE').contentWindow.document.body.innerHTML = file;
		window.close();
	}
}

function openFile(){
	if (document.getElementById('URL').value != ''){
 		document.getElementById('loader').contentWindow.document.location.href = document.getElementById('URL').value;
	}
}
</script>
<link href="<% = strRTECSSfile %>default_style.css" rel="stylesheet" type="text/css" />
</head>
<body style="margin:0px;" OnLoad="self.focus(); document.getElementById('Submit').disabled = true;">
<table width="100%"  border="0" cellpadding="3" cellspacing="0" class="RTEtableTopRow">
    <tr class="RTEtableTopRow">
      <td colspan="2" width="57%"><h1><% = strTxtOpen %></h1></td>
    </tr>
    <tr>
      <td colspan="2" class="RTEtableRow"><table width="100%" border="0" align="center" cellpadding="1" cellspacing="0">
        <tr>
          <td width="58%" valign="top">
            <table width="100%" border="0" cellspacing="0" cellpadding="2">
              <tr>
                <td width="51%" class="text"><% = strTxtPath %>: <span id="path"><% = "/" & strOpenFilePath %></span></td>
                <td width="49%" class="text"><% = strTxtFileName %>: <span id="fileName"></span>
                  <input name="URL" type="hidden" id="URL" />
                </td>
              </tr>
              <tr>
                <td class="text"><% = strTxtFileName %>:<iframe src="RTE_file_browser.asp?look=open" id="fileWindow" width="98%" height="215px" style="border: #A5ACB2 1px solid;backgroundColor: #FFFFFF;"></iframe></td>
                <td class="text"><% = strTxtPreview %>:<iframe src="RTE_popup_link_preview.asp" id="prevWindow" width="98%" height="215px" style="border: #A5ACB2 1px solid;backgroundColor: #FFFFFF;"></iframe></td>
                </tr>
          </table></td>
        </tr>
      </table></td>
    </tr>
    <tr>
    <td class="RTEtableBottomRow" valign="top">&nbsp;<%

'***** START WARNING - REMOVAL OR MODIFICATION OF THIS CODE WILL VIOLATE THE LICENSE AGREEMENT ******
If blnAbout Then
	Response.Write("<span class=""text"" style=""font-size:10px""><a href=""http://www.richtexteditor.org"" target=""_blank"" style=""font-size:10px"">Web Wiz Rich Text Editor</a> version " & strRTEversion & "</span>")
End If 
'***** END WARNING - REMOVAL OR MODIFICATION OF THIS CODE WILL VIOLATE THE LICENSE AGREEMENT ******      
      
      %></td>
      <td align="right" class="RTEtableBottomRow"><iframe src="" id="loader" width="0px" height="0px" style="visibility:hidden; position: absolute; left: 0px; top: 0px;" onLoad="getFile();"></iframe>
          <input type="button" name="Submit" id="Submit" value="   <% = strTxtOpen %>   " onClick="openFile()">&nbsp;<input type="button" name="cancel" value=" <% = strTxtCancel %> " onClick="window.close()">
        <br /><br /></td>
    </tr>
</table>
</body>
</html>