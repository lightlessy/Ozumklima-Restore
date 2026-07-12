<%@ Language="VBScript" CodePage="65001" %>
<%
Response.Buffer = True
Response.CodePage = 65001
Response.Charset = "utf-8"
%>
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




'Set the response buffer to true as we maybe redirecting
Response.Buffer = True

'Make sure this page is not cached
Response.Expires = -1
Response.ExpiresAbsolute = Now() - 2
Response.AddHeader "pragma","no-cache"
Response.AddHeader "cache-control","private"
Response.CacheControl = "No-Store"


Dim strEditorContent 	'Holds the content from the editors text area




'Read in the data posted from the editor
strEditorContent = Request.Form("myTextarea")



'*******************************************************************
'*** 	Start - What can I do with the data from the editor??? 	****
'*******************************************************************


'In this eample we are just displaying the data in a web page, but you could enter it into a database

'You could place the data from the editor in the variable strEditorContent into a database here

'To findout how to save data to an MS Access database please see:-
' 	Connecting to an Access database - http://www.webwiz.co.uk/kb/asp_tutorials/connecting_to_a_database.asp
' 	Saving data to an Access database - http://www.webwiz.co.uk/kb/asp_tutorials/add_to_database.asp


'*******************************************************************
'*** 	End - What can I do with the data from the editor??? 	****
'*******************************************************************



'If there is nothing to preview then say so
If strEditorContent = "" OR IsNull(strEditorContent) OR (InStr(1, strEditorContent, "<br />", 1) = 1 AND Len(strEditorContent) = 8) Then
	strEditorContent = "<br /><br /><div align=""center"">" & strTxtThereIsNothingToShow & "</div><br /><br />"
End If

%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta name="copyright" content="Copyright (C) 2001-2012 Web Wiz" />
<title>Display Message</title>
<HTTP-EQUIV="PRAGMA" CONTENT="NO-CACHE">
<!--#include file="RTE_configuration/browser_page_encoding_inc.asp" -->

<!--//
/* *******************************************************
Software: Web Wiz Rich Text Editor(TM) ver. <% = strRTEversion %>
Info: http://www.richtexteditor.org
Copyright: (C)2001-2012 Web Wiz Ltd. All rights reserved
******************************************************* */
//-->

<link href="<% = strRTECSSfile %>default_style.css" rel="stylesheet" type="text/css" />
</head>
<body OnLoad="self.focus();">
<table width="100%" border="0" cellspacing="0" cellpadding="1" align="center" height="53">
  <tr>
    <td align="center" height="17"><h1><% = strTxtDisplayMessage %></h1></td>
  </tr>
</table>
<table width="98%" border="0" cellspacing="0" cellpadding="1" bgcolor="#E7E7E7" align="center">
 <tr>
  <td>
          <table width="100%" border="0" cellspacing="0" cellpadding="2" bgcolor="#FFFFFF" height="147">
          <tr>
           <td class="RTEtextarea" valign="top">
            <!--// /* Message body -->
            <% = strEditorContent %>
            <!-- Message body ''"" */ //-->
           </td>
          </tr>
         </table>
         </td>
       </tr>
      </table>
<div align="center">
<br />
<%
'***** START WARNING - REMOVAL OR MODIFICATION OF THIS CODE WILL VIOLATE THE LICENSE AGREEMENT ******
If blnAbout Then
	Response.Write("<span class=""text"" style=""font-size:10px""><a href=""http://www.richtexteditor.org"" target=""_blank"" style=""font-size:10px"">Web Wiz Rich Text Editor</a> version " & strRTEversion & "</span>")
	Response.Write("<br /><span class=""text"" style=""font-size:10px"">Copyright &copy;2001-2012 Web Wiz</span>")
End If
'***** END WARNING - REMOVAL OR MODIFICATION OF THIS CODE WILL VIOLATE THE LICENSE AGREEMENT ******
%>
</div>
</body>
</html>