
<!--#include file="functions/RTE_functions_common.asp" -->
<!--#include file="RTE_configuration/RTE_setup.asp" -->
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






'***** START WARNING - REMOVAL OR MODIFICATION OF THIS CODE WILL VIOLATE THE LICENSE AGREEMENT ******
Response.Write(vbCrLf & vbCrLf & "<!--//" & _
vbCrLf & "/* *******************************************************" & _
vbCrLf & "Software: Web Wiz Rich Text Editor(TM) ver. " & strRTEversion & "" & _
vbCrLf & "Info: http://www.richtexteditor.org" & _
vbCrLf & "Copyright: (C)2001-2012 Web Wiz Ltd. All rights reserved" & _
vbCrLf & "******************************************************* */" & _
vbCrLf & "//-->")
'***** END WARNING - REMOVAL OR MODIFICATION OF THIS CODE WILL VIOLATE THE LICENSE AGREEMENT ******



'Load default CSS and Javascript
Response.Write(vbCrLf & vbCrLf & "<link href=""" & strRTECSSfile & "RTE_toolbar_style.css"" rel=""stylesheet"" type=""text/css"" />")
Response.Write(vbCrLf & "<script language=""JavaScript"" src=""RTE_javascript_common.asp"" type=""text/javascript""></script>")


'If this is an RTE enabled web browser load in the RTE content
If RTEenabled() <> "false" Then
	
	'Load in Javascript for RTE browsers
	Response.Write(vbCrLf & "<script language=""JavaScript"" src=""RTE_javascript.asp?textArea=" & Server.URLEncode(strTextAreaName) & """ type=""text/javascript""></script>")

	

'If this is not an RTE enabled web browser load in the NON-RTE content
Else

	'Load in Javascript for non-RTE browsers
	Response.Write(vbCrLf & "<script language=""JavaScript"" src=""non_RTE_javascript.asp?textArea=" & Server.URLEncode(strTextAreaName) & """ type=""text/javascript""></script>")

	'Display an message that the user is using the non-RTE editor
	Response.Write(vbCrLf & "<script language=""JavaScript"">")
	Response.Write(vbCrLf & "	alert('Your web browser is not Rich Text Enabled.\nYou are using the standard text box area.');")
	Response.Write(vbCrLf & "</script>")
End If

'Create Toolbar and Iframes
Response.Write(vbCrLf & "<script language=""JavaScript"">WebWizRTEtoolbar('" & strFormName & "');</script>")

%>