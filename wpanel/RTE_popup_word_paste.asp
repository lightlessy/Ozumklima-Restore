<% @ Language=VBScript %>
<% Option Explicit %>  



<!--#include file="RTE_configuration/RTE_setup.asp" -->
<!--#include file="language_files/RTE_language_file_inc.asp" -->
<!--#include file="functions/RTE_functions_common.asp" -->
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

%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta charset="utf-8">
<title>Paste from Word</title>
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


'If this is Gecko based browser or Opera link to JS code for Gecko
If RTEenabled = "Gecko" OR RTEenabled = "opera" Then Response.Write(vbCrLf & "<script language=""JavaScript"" src=""RTE_javascript_gecko.asp"" type=""text/javascript""></script>")
	
%>
<script language="JavaScript">
//function to initialise paste textarea
function initialise(){

	//create iframe page content
	var editor = document.getElementById('pasteWin').contentWindow.document;
	var iframeContent;
	iframeContent  = '<html>\n';
	iframeContent += '<head>\n';
	iframeContent += '<style> html,body{border:0px;background-color:#FFFFFF;}</style>\n';
	iframeContent += '</head>\n';
	iframeContent += '<body leftmargin="1" topmargin="1" marginwidth="1" marginheight="1">\n';
	iframeContent += '</body>\n';
	iframeContent += '</html>';

	editor.open();
	editor.write(iframeContent);
	editor.close();

	function initIframe() {
		editor.designMode = 'on';
	};
	setTimeout(initIframe, 100);
	self.focus(); 
}


//Get pasted word document
function pasteWordDoc(){

	//Read in the word doc
	pastedDoc = document.getElementById('pasteWin').contentWindow.document.body.innerHTML;
	
	//Run through Word Tidy function
	pastedDoc = WWRTEwordTidy(pastedDoc);
	
	if (pastedDoc.indexOf('<br>') > -1 && pastedDoc.length == 8) pastedDoc = '';
	
	//Place in main editor
	editor = window.opener.document.getElementById('WebWizRTE');	
<%
'If this is Mozilla or Opera then we need to call insertElementPosition to find where to place the doc
If RTEenabled = "Gecko" OR RTEenabled = "opera" Then 
%>     		
     	docHTML = editor.contentWindow.document.createElement('wordTidy');
     	docHTML.innerHTML = pastedDoc;
	try{
		insertElementPosition(editor.contentWindow, docHTML);
	}catch(exception){
		alert('<% = strTxtErrorInsertingObject %>');
		editor.contentWindow.focus();
	}
<%	
'Else this is IE so placing the doc is simpler
Else
%>
	try{
		editor.contentWindow.focus();
		editor.contentWindow.document.selection.createRange().pasteHTML(pastedDoc);
	}catch(exception){
		alert('<% = strTxtErrorInsertingObject %>');
		editor.contentWindow.focus();
	}
<%
End If

'Set focus
'If Opera change the focus method
If RTEenabled = "opera" Then
		
	Response.Write(vbCrLf & "	editor.focus();")
Else
	Response.Write(vbCrLf & "	editor.contentWindow.focus();")
End If
%>
	window.close();
}


//Clean word HTML using WordTidy(TM) Technology
function WWRTEwordTidy(doc){

	//Delete all SPAN tags
	doc = doc.replace(/<\/?SPAN[^>]*>/gi, '')
	
	//Delete all FONT tags
	.replace(/<\/?FONT[^>]*>/gi, '')
	
	//Delete Class attributes
	.replace(/<(\w[^>]*) class=([^ |>]*)([^>]*)/gi, '<$1$3')
	
	//Delete Style attributes
	.replace(/<(\w[^>]*) style='([^']*)'([^>]*)/gi, '<$1$3')

	//Delete Lang attributes
	.replace(/<(\w[^>]*) lang=([^ |>]*)([^>]*)/gi, '<$1$3')
	
	//Delete XML elements and declarations
	.replace(/<\\?\?xml[^>]*>/gi, '')
	
	//Delete local file links
	.replace(/<link rel=[^>]*>/gi,'')
	
	//Delete meta tags
	.replace(/<meta [^>]*>/gi,'')
	
	//Delete style
	.replace(/<\/?style[^>]*>/gi, '')
	
	//Delete Tags with XML namespace declarations: <o:p></o:p>
	.replace(/<\/?\w+:[^>]*>/gi, '')
	
	//Delete the &nbsp;
	.replace(/&nbsp;/, ' ')
	
	//Delete the MARGIN: 0cm 0cm 0pt; IE puts in when pasting from Word
	.replace(/MARGIN: 0cm 0cm 0pt;/gi, '')
	
	//Clean up tags
	.replace(/<B [^>]*>/gi,'<b>')
	.replace(/<I [^>]*>/gi,'<i>')
	.replace(/<LI [^>]*>/gi,'<li>')
	.replace(/<UL [^>]*>/gi,'<ul>')
	
	//Replace outdated tags
	.replace(/<B>/gi,'<strong>')
	.replace(/<\/B>/gi,'</strong>')
	.replace(/<I>/gi,'<em>')
	.replace(/<\/I>/gi,'</em>')
	
	//Delete empty tags
	.replace(/<strong><\/strong>/gi,'')
	.replace(/<strong> <\/strong>/gi,'')
	.replace(/<em><\/em>/gi,'')
	.replace(/<em> <\/em>/gi,'')
	
	//Replace <P> with <DIV>
	.replace(/<P/gi, '<div')
	.replace(/<\/P>/gi, '</div>')
	
	//Replace single smartquotes ''
	.replace(/['']/g, "'")
	//Replace double smartquotes ""
	.replace(/[""]/g, '"')
	
	return doc;
}
</script>
<link href="<% = strRTECSSfile %>default_style.css" rel="stylesheet" type="text/css" />
</head>
<body style="margin:0px;" OnLoad="initialise()">
<table width="100%"  border="0" cellpadding="3" cellspacing="0" class="RTEtableTopRow">
    <tr class="RTEtableTopRow">
      <td colspan="2" width="57%"><h1><% = strTxtPasteFromWord %></h1></td>
    </tr>
    <tr>
      <td colspan="2" class="RTEtableRow"><table width="100%" border="0" align="center" cellpadding="1" cellspacing="0">
        <tr>
          <td width="58%" valign="top">
            <table width="100%" border="0" cellspacing="0" cellpadding="2">
              <tr>
                <td width="51%" class="text"><% = strTxtPasteFromWordDialog %></td>
              </tr>
              <tr>
                <td class="text"><iframe id="pasteWin" width="100%" height="180px" style="border: #A5ACB2 1px solid;backgroundColor: #FFFFFF;"></iframe></td>
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
      <td align="right" class="RTEtableBottomRow">
          <input type="button" name="Submit" id="Submit" value="   <% = strTxtOK %>   " onClick="pasteWordDoc()">&nbsp;<input type="button" name="cancel" value=" <% = strTxtCancel %> " onClick="window.close()">
        <br /><br /></td>
    </tr>
</table>
</body>
</html>