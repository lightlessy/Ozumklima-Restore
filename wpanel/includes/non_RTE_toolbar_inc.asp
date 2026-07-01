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



'***********************************
'****	non-RTE Toolbar 1	****
'***********************************


'File toolbar
'------------
If blnNew Then 
	Response.Write(vbCrLf & "	document.writeln('<img src=""" & strRTEImagePath & "toolbar_separator.gif"" align=""absmiddle"" />")
	Response.Write("<img src=""" & strRTEImagePath & "post_button_new.gif"" align=""absmiddle"" title=""" & strTxtNewBlankDoc & """ onClick=""clearWebWizRTE()"" class=""WebWizRTEbutton"" onmouseover=""overIcon(this)"" onmouseout=""outIcon(this)"" />")
	Response.Write("');")
End If


'Preview, print, spell toolbar
'------------
If blnPreview Then 
	Response.Write(vbCrLf & "	document.writeln('<img src=""" & strRTEImagePath & "toolbar_separator.gif"" align=""absmiddle"" />")
	'Button Pop up Preview
	Response.Write("<img src=""" & strRTEImagePath & "post_button_preview.gif"" align=""absmiddle"" border=""0"" title=""" & strTxtPreview & """ onClick=""document.getElementById(\'pre\').value=document.getElementById(textArea).value; OpenPreviewWindow(document.' + formName + ');"" class=""WebWizRTEbutton"" onmouseover=""overIcon(this)"" onmouseout=""outIcon(this)"" />")
	Response.Write("<input type=""hidden"" name=""pre"" id=""pre"" value="""" />")
	Response.Write("');")
End If


'insert toolbar
'------------
If blnAdvAdddHyperlink OR blnAddHyperlink Or blnAttachments OR blnAdvAddImage OR blnAddImage OR blnImageUpload Then Response.Write(vbCrLf & "	document.writeln('<img src=""" & strRTEImagePath & "toolbar_separator.gif"" align=""absmiddle"" />")
'If hyperlink is enabled
If blnAdvAdddHyperlink OR blnAddHyperlink Then 
	Response.Write("<img src=""" & strRTEImagePath & "post_button_hyperlink.gif"" align=""absmiddle"" border=""0"" onClick=""AddCode(\'URL\')"" title=""" & strTxtAddHyperlink & """ class=""WebWizRTEbutton"" onmouseover=""overIcon(this)"" onmouseout=""outIcon(this)"" />")
	Response.Write("<img src=""" & strRTEImagePath & "post_button_email.gif"" align=""absmiddle"" border=""0"" onClick=""AddCode(\'EMAIL\')"" title=""" & strTxtAddEmailLink & """ class=""WebWizRTEbutton"" onmouseover=""overIcon(this)"" onmouseout=""outIcon(this)"" />")
End If
If blnAttachments Then Response.Write("<img src=""" & strRTEImagePath & "post_button_file_upload.gif"" align=""absmiddle"" onClick=""winOpener(\'non_RTE_upload_files.asp?textArea=\'+textArea,\'files\',0,1,400,163)"" title=""" & strTxtFileUpload & """ border=""0"" class=""WebWizRTEbutton"" onmouseover=""overIcon(this)"" onmouseout=""outIcon(this)"" />")
'images
If blnAdvAddImage OR blnAddImage Then Response.Write("<img src=""" & strRTEImagePath & "post_button_image.gif"" align=""absmiddle"" border=""0"" onClick=""AddCode(\'IMG\')"" title=""" & strTxtAddImage & """ class=""WebWizRTEbutton"" onmouseover=""overIcon(this)"" onmouseout=""outIcon(this)"" />")
'If image uploading is allowed have an image upload button
If blnImageUpload Then Response.Write("<img src=""" & strRTEImagePath & "post_button_image_upload.gif"" align=""absmiddle"" border=""0"" onClick=""winOpener(\'non_RTE_upload_images.asp?textArea=\'+textArea,\'images\',0,1,400,150)"" title=""" & strTxtImageUpload & """ class=""WebWizRTEbutton"" onmouseover=""overIcon(this)"" onmouseout=""outIcon(this)"" />")
If blnAdvAdddHyperlink OR blnAddHyperlink Or blnAttachments OR blnAdvAddImage OR blnAddImage OR blnImageUpload Then Response.Write("');")


'List type and indent toolbar
'------------
If blnOrderList OR blnUnOrderList OR blnOutdent OR blnIndent Then Response.Write(vbCrLf & "	document.writeln('<img src=""" & strRTEImagePath & "toolbar_separator.gif"" align=""absmiddle"" />")
If blnOrderList OR blnUnOrderList Then Response.Write("<img src=""" & strRTEImagePath & "post_button_or_list.gif"" align=""absmiddle"" border=""0"" onClick=""AddCode(\'LIST\', \'\')"" title=""" & strTxtList & """ class=""WebWizRTEbutton"" onmouseover=""overIcon(this)"" onmouseout=""outIcon(this)"" />")
If blnOutdent Then Response.Write("<img src=""" & strRTEImagePath & "post_button_outdent.gif"" align=""absmiddle"" onClick=""FormatText(\'Outdent\', \'\')"" title=""" & strTxtOutdent & """ class=""WebWizRTEbutton"" onmouseover=""overIcon(this)"" onmouseout=""outIcon(this)"" />")
If blnIndent Then Response.Write("<img src=""" & strRTEImagePath & "post_button_indent.gif"" align=""absmiddle"" border=""0"" onClick=""AddCode(\'INDENT\', \'\')"" title=""" & strTxtIndent & """ class=""WebWizRTEbutton"" onmouseover=""overIcon(this)"" onmouseout=""outIcon(this)"" />") 
If blnOrderList OR blnUnOrderList OR blnOutdent OR blnIndent Then Response.Write("');")


Response.Write(vbCrLf & "	document.writeln('<br />');")




'***********************************
'****	non-RTE Toolbar 2	****
'***********************************

		
'Toolbar buttons


'Font type toolbar
'------------
If blnFontStyle OR blnFontType OR blnFontSize OR blnTextColour Then Response.Write(vbCrLf & "	document.writeln('<img src=""" & strRTEImagePath & "toolbar_separator.gif"" align=""absmiddle"" />")
'Font Type
If blnFontType  Then Response.Write(" <select name=""selectFont"" onChange=""FontCode(selectFont.options[selectFont.selectedIndex].value, \'FONT\');selectFont.options[0].selected = true;""><option selected>" & strTxtFontTypes & "</option><option value=\'font=""Arial""\'>Arial</option><option value=\'font=""Courier""\'>Courier New</option><option value=\'font=""Times""\'>Times</option><option value=\'font=""Verdana""\'>Verdana</option></select>")
'Font Size
If blnFontSize  Then Response.Write(" <select name=""selectSize"" onChange=""FontCode(selectSize.options[selectSize.selectedIndex].value, \'SIZE\');selectSize.options[0].selected = true;""><option selected>" & strTxtFontSizes & "</option><option value=\'font size=""1""\'>1</option><option value=\'font size=""2""\'>2</option><option value=\'font size=""3""\'>3</option><option value=\'font size=""4""\'>4</option><option value=\'font size=""5""\'>5</option><option value=\'font size=""6""\'>6</option></select>")
'Font colour
If blnTextColour Then Response.Write(" <select name=""selectColour"" onChange=""FontCode(selectColour.options[selectColour.selectedIndex].value, \'COLOR\');selectColour.options[0].selected = true;""><option value=""0"" selected>" & strTxtFontColour & "</option><option value=\'font color=""black""\'>" & strTxtBlack & "</option><option value=\'font color=""white""\'>" & strTxtWhite & "</option><option value=\'font color=""blue""\'>" & strTxtBlue & "</option><option value=\'font color=""red""\'>" & strTxtRed & "</option><option value=\'font color=""green""\'>" & strTxtGreen & "</option><option value=\'font color=""yellow""\'>" & strTxtYellow & "</option><option value=\'font color=""orange""\'>" & strTxtOrange & "</option><option value=\'font color=""brown""\'>" & strTxtBrown & "</option><option value=\'font color=""magenta""\'>" & strTxtMagenta & "</option><option value=\'font color=""cyan""\'>" & strTxtCyan & "</option><option value=\'font color=""lime green""\'>" & strTxtLimeGreen & "</option></select>")
If blnFontStyle OR blnFontType OR blnFontSize OR blnTextColour Then Response.Write("');")


'Font style toolbar
'------------
If blnBold OR blnItalic OR blnUnderline Then Response.Write(vbCrLf & "	document.writeln('<img src=""" & strRTEImagePath & "toolbar_separator.gif"" align=""absmiddle"" />")
If blnBold Then Response.Write("<img src=""" & strRTEImagePath & "post_button_bold.gif"" align=""absmiddle"" title=""" & strTxtBold & """ onClick=""AddMessageCode(\'B\',\'" & strTxtEnterBoldText & "\', \'\')"" class=""WebWizRTEbutton"" onmouseover=""overIcon(this)"" onmouseout=""outIcon(this)"" />")
If blnItalic Then Response.Write("<img src=""" & strRTEImagePath & "post_button_italic.gif""  align=""absmiddle"" title=""" & strTxtItalic & """ onClick=""AddMessageCode(\'I\',\'" & strTxtEnterItalicText & "\', \'\')"" class=""WebWizRTEbutton"" onmouseover=""overIcon(this)"" onmouseout=""outIcon(this)"" />")
If blnUnderline Then Response.Write("<img src=""" & strRTEImagePath & "post_button_underline.gif"" align=""absmiddle"" title=""" & strTxtUnderline & """ onClick=""AddMessageCode(\'U\',\'" & strTxtEnterUnderlineText & "\', \'\')"" class=""WebWizRTEbutton"" onmouseover=""overIcon(this)"" onmouseout=""outIcon(this)"" />")            
If blnBold OR blnItalic OR blnUnderline Then Response.Write("');")


'Font block format toolbar
'------------
If blnCentre Then 
	Response.Write(vbCrLf & "	document.writeln('<img src=""" & strRTEImagePath & "toolbar_separator.gif"" align=""absmiddle"" />")
	Response.Write("<img src=""" & strRTEImagePath & "post_button_centre.gif"" align=""absmiddle"" border=""0"" onClick=""AddMessageCode(\'center\',\'" & strTxtEnterCentredText & "\', \'\')"" title=""" & strTxtCentrejustify & """ class=""WebWizRTEbutton"" onmouseover=""overIcon(this)"" onmouseout=""outIcon(this)"" />") 
	Response.Write("');")
End If


'About toolbar
'------------
If blnAbout Then 
	Response.Write(vbCrLf & "	document.writeln('<img src=""" & strRTEImagePath & "toolbar_separator.gif"" align=""absmiddle"" />")
	Response.Write("<img src=""" & strRTEImagePath & "post_button_about.gif"" align=""absmiddle"" border=""0"" onClick=""winOpener(\'RTE_popup_about.asp\',\'about\',0,0,400,187)"" title=""" & strTxtAboutRichTextEditor & """ class=""WebWizRTEbutton"" onmouseover=""overIcon(this)"" onmouseout=""outIcon(this)"" />")
	Response.Write("');")
End If

'Mode
Response.Write(vbCrLf & "	document.writeln('&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<select name=""selectMode"" onChange=""PromptMode(this)""><option value=""1"" selected>" & strTxtPrompt & " " & strTxtMode & "</option><option value=""0"">" & strTxtBasic & " " & strTxtMode & "</option></select>');")
%>