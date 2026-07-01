
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

Dim strTextAreaName

strTextAreaName = removeAllTags(Request.QueryString("textArea"))

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

//text area name
var textArea = '<% = strTextAreaName %>';

//Have the propmt box turned on by default
var promptOn = true;


//initialise area
function initialiseWebWizRTE(){
	//make toolbar the same size
	document.getElementById('toolbar').width = parseInt(document.getElementById(textArea).offsetWidth) + 2 + 'px';
}

//Create toolbar
function WebWizRTEtoolbar(formName){

	document.writeln('<table id="toolbar" width="680" border="0" cellspacing="0" cellpadding="1" class="RTEtoolbar"');
	document.writeln(' <tr>');
	document.writeln('  <td>');
	<!--#include file="includes/non_RTE_toolbar_inc.asp" -->
	document.writeln('  </td>');
	document.writeln(' </tr>');
	document.writeln('</table>');

}

//Function to turn on or off the prompt box
function PromptMode(selectMode){
	if (selectMode.options[selectMode.selectedIndex].value == 0){
		promptOn = false;
	}
	else{
		promptOn = true;
	}
}

//function to clear area
function clearWebWizRTE(){
 	if (window.confirm('<% = strResetWarningEditorConfirm %>')){
	 	document.getElementById(textArea).value='';
	}
}

// Function to add the code for bold italic centre and underline, to the message
function AddMessageCode(code, promptText, InsertText){

	if (code != ''){
		if (promptOn == true){
			insertCode = prompt(promptText + '\n<' + code + '>xxx</' + code + '>', InsertText);
				if ((insertCode != null) && (insertCode != '')){
					document.getElementById(textArea).value += '<' + code + '>' + insertCode + '</' + code + '>';
				}
		}
		else{
			document.getElementById(textArea).value += '<' + code + '></' + code + '>';
		}
	}
				
	document.getElementById(textArea).focus();
}

// Function to add the font colours, sizes, type to the message
function FontCode(code, endCode) {

	if (code != ''){
		if (promptOn == true){
			insertCode = prompt('<% = strTxtEnterTextYouWouldLikeIn %> ' + code + '\n<' + code + '>xxx</font>', '');
				if ((insertCode != null) && (insertCode != '')){
					document.getElementById(textArea).value += '<' + code + '>' + insertCode + '</font>';
				}
		}
		else{
			document.getElementById(textArea).value += '<' + code + '></font>';
		}
	}	
	document.getElementById(textArea).focus();
}



//Function to add the URL, indent, list, and Email code to the message
function AddCode(code){

	//For the URL code
	if ((code != '') && (code == 'URL')){
		if (promptOn == true){
			insertText = prompt('<% = strTxtEnterHyperlinkText %>', '');
				
				if ((insertText != null) && (insertText != '') && (code == 'URL')){
					insertCode = prompt('<% = strTxtEnterHeperlinkURL %>', 'http://');
						
						if ((insertCode != null) && (insertCode != '') && (insertCode != 'http://')){					
							document.getElementById(textArea).value += '<a href="' + insertCode + '">' + insertText + '</a>';
						}
				}
		}
		else {
			document.getElementById(textArea).value += '<a href=""></a>';
		}
	}
	
	
	//For the email code
	if ((code != '') && (code == 'EMAIL')){
		if (promptOn == true){
			insertText = prompt('<% = strTxtEnterEmailText %>', '');
				
				if ((insertText != null) && (insertText != '')){
					insertCode = prompt('<% = strTxtEnterEmailMailto %>', '');
						
						if ((insertCode != null) && (insertCode != '')){					
						document.getElementById(textArea).value += '<a href="mailto:' + insertCode + '">' + insertText + '</a>';
					}
				}
		}
		else {
			document.getElementById(textArea).value += '<a href="mailto:"></a>';
		}
	}
	
	//For the image code
	if ((code != '') && (code == 'IMG')){
		if (promptOn == true){	
			insertCode = prompt('<% = strTxtEnterImageURL %>', 'http://');
						
				if ((insertCode != null) && (insertCode != '')){					
				document.getElementById(textArea).value += '<img src="' + insertCode + '" />';
			}			
		}
		else {
			document.getElementById("message").value += '<img src="" />';
		}
	}
	
	//For the list code
	if ((code != '') && (code == 'LIST')){
		if (promptOn == true){
			listType = prompt('<% = strTxtEnterTypeOfList %> \n<% = strTxtEnterEnter %> \'1\' <% = strTxtEnterNumOrBlankList %>', '');
			
			while ((listType != null) && (listType != '') && (listType != '1')) {
				listType = prompt('<% = strTxtEnterListError %> \'1\' <% = strTxtEnterNumOrBlankList %>','');               
			}
			
			if (listType != null){			
				var listItem = '1';
				var insertCode = '';
				
				while ((listItem != '') && (listItem != null)) {
					listItem = prompt('<% = strEnterLeaveBlankForEndList %>',''); 
					if (listItem != '') {             
						insertCode += '<li>' + listItem + '</li>'; 
					}                   
				} 
				
				if (listType == '') {
					document.getElementById(textArea).value += '<ul>' + insertCode + '</ul>';
				} else {
					document.getElementById(textArea).value += '<ol>' + insertCode + '</ol>';
				} 
				
			}
		}
		else{
			document.getElementById(textArea).value += '<ul><li> </li><li> </li><li> </li></ul>';
		}			
	}
	
	
	//For the indent
	if ((code != '') && (code == 'INDENT')){
						
			document.getElementById(textArea).value += '      ';				
	}
				
	document.getElementById(textArea).focus();
}