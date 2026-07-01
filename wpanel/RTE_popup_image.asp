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




'Dimension veriables
Dim strImageURL
Dim strImageAltText
Dim strAlign
Dim intBorder
Dim lngHorizontal
Dim lngVerical
Dim intWidth
Dim intHeight


'If this a post back read in the form elements
If Request.Form("URL") <> "http://" AND Request.Form("URL") <> "" AND Request.Form("postBack") Then
	
	'Initilise variable
	intBorder = 0
	
	'Get form elements
	strImageURL = Request.Form("URL")
	strImageAltText = Request.Form("Alt")
	strAlign = Request.Form("align")
	If isNumeric(Request.Form("intBorder")) Then intBorder = Request.Form("border")
	If isNumeric(Request.Form("hoz")) Then lngHorizontal = LngC(Request.Form("hoz"))
	If isNumeric(Request.Form("vert")) Then lngVerical = LngC(Request.Form("vert"))
	If isNumeric(Request.Form("width")) Then intWidth = LngC(Request.Form("width"))
	If isNumeric(Request.Form("height")) Then intHeight = LngC(Request.Form("height"))
	
	'Escape characters that will course a crash
	strImageURL = Replace(strImageURL, "'", "\'", 1, -1, 1)
	strImageURL = Replace(strImageURL, """", "\""", 1, -1, 1)
	strImageAltText = Replace(strImageAltText, "'", "\'", 1, -1, 1)
	strImageAltText = Replace(strImageAltText, """", "\""", 1, -1, 1)
End If

%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<title>Image Properties</title>
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
<%
'If this a post back write javascript
If Request.Form("URL") <> "http://" AND Request.Form("URL") <> "" AND Request.Form("postBack") Then
	
	Response.Write(vbCrLf & vbCrLf & "	editor = window.opener.document.getElementById('WebWizRTE');")
	
	'Tell that we are an image
	Response.Write(vbCrLf & vbCrLf & "	img = editor.contentWindow.document.createElement('img');")
	
	'Set image attributes
	Response.Write(vbCrLf & vbCrLf & "	img.setAttribute('src', '" & strImageURL & "');")
	Response.Write(vbCrLf & "	img.setAttribute('border', '" & intBorder & "');")
	If strImageAltText <> "" Then Response.Write(vbCrLf & "	img.setAttribute('alt', '" & strImageAltText & "');")
	If lngHorizontal <> "" Then Response.Write(vbCrLf & "	img.setAttribute('hspace', '" & lngHorizontal & "');")
	If intWidth <> "" Then Response.Write(vbCrLf & "	img.setAttribute('width', '" & intWidth & "');")
	If intHeight <> "" Then Response.Write(vbCrLf & "	img.setAttribute('height', '" & intHeight & "');")
	If lngVerical <> "" Then Response.Write(vbCrLf & "	img.setAttribute('vspace', '" & lngVerical & "');")
	If strAlign <> "" Then Response.Write(vbCrLf & "	img.setAttribute('align', '" & strAlign & "');")
	 
     
     	'If this is Mozilla or Opera then we need to call insertElementPosition to find where to place the image
     	If RTEenabled = "Gecko" OR RTEenabled = "opera" Then 
		
		Response.Write(vbCrLf & vbCrLf & "	try{" & _
					vbCrLf & "		insertElementPosition(editor.contentWindow, img);" & _
					vbCrLf & "	}catch(exception){" & _
					vbCrLf & "		alert('" & strTxtErrorInsertingObject & "');" & _
					vbCrLf & "		editor.contentWindow.focus();" & _
					vbCrLf & "	}")
	
	'Else this is IE so placing the image is simpler
	Else
		Response.Write(vbCrLf & vbCrLf & "	try{" & _
					vbCrLf & "		editor.contentWindow.focus();" & _
					vbCrLf & "		editor.contentWindow.document.selection.createRange().pasteHTML(img.outerHTML);" & _
					vbCrLf & "	}catch(exception){" & _
					vbCrLf & "		alert('" & strTxtErrorInsertingObject & "');" & _
					vbCrLf & "		editor.contentWindow.focus();" & _
					vbCrLf & "	}")
	End If
		
	'Set focus
	'If Opera change the focus method
	If RTEenabled = "opera" Then
		
		Response.Write(vbCrLf & "	editor.focus();")
	Else
		Response.Write(vbCrLf & "	editor.contentWindow.focus();")
	End If
		
	'Close window
	Response.Write(vbCrLf & "	window.close();")
End If

%>

//Function to preview image
function getImage(URL){
	document.getElementById("prevWindow").contentWindow.document.getElementById("prevFile").src = URL
}

//Function to change image properties
function changeImage(){
	document.getElementById("prevWindow").contentWindow.document.getElementById("prevFile").alt = document.getElementById('Alt').value
	document.getElementById("prevWindow").contentWindow.document.getElementById("prevFile").align = document.getElementById('align').value
	document.getElementById("prevWindow").contentWindow.document.getElementById("prevFile").border = document.getElementById('border').value
	document.getElementById("prevWindow").contentWindow.document.getElementById("prevFile").hspace = document.getElementById('hoz').value
	document.getElementById("prevWindow").contentWindow.document.getElementById("prevFile").vspace = document.getElementById('vert').value
	if (document.getElementById('width').value!=''){
		document.getElementById("prevWindow").contentWindow.document.getElementById("prevFile").width = document.getElementById('width').value
	}
	if (document.getElementById('height').value!=''){
		document.getElementById("prevWindow").contentWindow.document.getElementById("prevFile").height = document.getElementById('height').value
	}
}
</script>
<link href="<% = strRTECSSfile %>default_style.css" rel="stylesheet" type="text/css" />
</head>
<body style="margin:0px;" OnLoad="self.focus(); document.getElementById('Submit').disabled=true;">
<table width="100%"  border="0" cellpadding="3" cellspacing="0" class="RTEtableTopRow">
  <form method="post" name="frmImageInsrt">
    <tr class="RTEtableTopRow">
      <td colspan="2"><h1><% = strTxtImageProperties %></h1></td>
    </tr>
    <tr>
      <td colspan="2" class="RTEtableRow"><table width="100%" border="0" align="center" cellpadding="1" cellspacing="0">
        <tr>
          <td width="100%">
            <table width="100%" border="0" cellspacing="0" cellpadding="2">
              <tr>
                <td width="17%" align="right" class="text"><% = strTxtImageURL %>:</td>
                <td colspan="5"><input name="URL" type="text" id="URL" size="40" value="http://" onChange="document.getElementById('Submit').disabled=false;" onkeypress="document.getElementById('Submit').disabled=false;">
                  <input name="preview" type="button" id="preview" value="<% = strTxtPreview %>" onClick="getImage(document.getElementById('URL').value)">
                </td>
              </tr>
              <tr>
                <td align="right" class="text"><% = strTxtAlternativeText %>:</td>
                <td colspan="5"><input name="Alt" type="text" id="Alt" size="40" onBlur="changeImage()"></td>
              </tr>
              <tr>
                <td align="right" class="text"><% = strTxtWidth %>:</td>
                <td width="5%"><input name="width" type="text" id="width" size="4" maxlength="3" onkeyup="changeImage()" autocomplete="off" /></td>
                <td width="22%" align="right" class="text"><% = strTxtHorizontal %>:</td>
                <td width="6%"><input name="hoz" type="text" id="hoz" size="4" maxlength="3" onkeyup="changeImage()" autocomplete="off" /></td>
                <td width="20%" align="right" class="text"><% = strTxtAlignment %>:</td>
                <td width="30%"><select size="1" name="align" id="align" onChange="changeImage()">
                  <option value="" selected >Default</option>
                  <option value="left">Left</option>
                  <option value="right">Right</option>
                  <option value="texttop">Texttop</option>
                  <option value="absmiddle">Absmiddle</option>
                  <option value="baseline">Baseline</option>
                  <option value="absbottom">Absbottom</option>
                  <option value="bottom">Bottom</option>
                  <option value="middle">Middle</option>
                  <option value="top">Top</option>
                </select></td>
              </tr>
              <tr>
                <td align="right" class="text"><% = strTxtHeight %>:</td>
                <td><input name="height" type="text" id="height" size="4" maxlength="3" onkeyup="changeImage()" autocomplete="off" /></td>
                <td align="right" class="text"><% = strTxtVertical %>:</td>
                <td><input name="vert" type="text" id="vert" size="4" maxlength="3" onkeyup="changeImage()" autocomplete="off" /></td>
                <td align="right" class="text"><% = strTxtBorder %>:</td>
                <td><input name="border" type="text" id="border" size="4" maxlength="2" onKeyUp="changeImage()" autocomplete="off" /></td>
              </tr>
              <tr>
                <td align="right" valign="top" class="text"><% = strTxtPreview %>:</td>
                <td colspan="5"><iframe src="RTE_popup_image_preview.asp" id="prevWindow" width="98%" height="180px" style="border: #A5ACB2 1px solid;backgroundColor: #FFFFFF;"></iframe></td>
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
      <td align="right" class="RTEtableBottomRow"><input type="hidden" name="postBack" value="true">
          <input type="submit" id="Submit" name="Submit" value="     <% = strTxtOK %>     ">&nbsp;<input type="button" name="cancel" value=" <% = strTxtCancel %> " onClick="window.close()">
        <br /><br /></td>
    </tr>
  </form>
</table>
</body>
</html>
