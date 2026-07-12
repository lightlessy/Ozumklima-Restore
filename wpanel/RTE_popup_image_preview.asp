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



Response.AddHeader "pragma","cache"
Response.AddHeader "cache-control","public"
Response.CacheControl = "Public"

%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<title>No Preview</title>
<!--#include file="RTE_configuration/browser_page_encoding_inc.asp" -->

<%
'***** START WARNING - REMOVAL OR MODIFICATION OF THIS CODE WILL VIOLATE THE LICENSE AGREEMENT ******
Response.Write(vbCrLf & "<!--//" & _
vbCrLf & "/* *******************************************************" & _
vbCrLf & "Software: Web Wiz Rich Text Editor " & _
vbCrLf & "Info: http://www.richtexteditor.org" & _
vbCrLf & "Copyright: (C)2001-2012 Web Wiz Ltd. All rights reserved" & _
vbCrLf & "******************************************************* */" & _
vbCrLf & "//-->")
'***** END WARNING - REMOVAL OR MODIFICATION OF THIS CODE WILL VIOLATE THE LICENSE AGREEMENT ******
%>

<script language="JavaScript">

//function to upadte image properties
function imageProperties(oImage){

	if (document.getElementById('prevFile').width != 1 && document.getElementById('prevFile').height != 1){
		window.parent.document.getElementById('width').value = document.getElementById('prevFile').width
		window.parent.document.getElementById('height').value = document.getElementById('prevFile').height
	}
}

</script>
<style type="text/css">
<!--
.text {
	font-family: Arial, Helvetica, sans-serif;
	font-size: 11px;
	color: #000000;
}
html,body {
	border: 0px;
}
-->
</style>
</head>
<body bgcolor="#FFFFFF" leftmargin="2" topmargin="2" marginwidth="2" marginheight="2">
<img src="<% = strRTEImagePath %>clear_pixel.gif" id="prevFile" onError="alert('<% = strTxtErrorLoadingPreview %>')" onLoad="imageProperties(this)">
</body>
</html>