<!--#include file="RTE_license.asp" -->
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

Dim blnAbout
Dim strDisplayLicense


'***** START WARNING - REMOVAL OR MODIFICATION OF THIS CODE WILL VIOLATE THE LICENSE AGREEMENT ******
Const strRTEversion = "4.09"
blnAbout = rteAbout()
If Request.QueryString("about") Then Call rteInfo()
'***** END WARNING - REMOVAL OR MODIFICATION OF THIS CODE WILL VIOLATE THE LICENSE AGREEMENT ******   



'The following enables and disables functions of the Rich Text Editor

'Enable and dsiable basic functions of the editor change the following to true of false
'***************************************************************************
Const blnNew = true
Const blnBold = true
Const blnUnderline = true
Const blnItalic = true
Const blnFontStyle = true
Const blnFontType = true
Const blnFontSize = true
Const blnTextColour = true
Const blnTextBackgroundColour = true
Const blnCut = true
Const blnCopy = true
Const blnPaste = true
Const blnWordPaste = true
Const blnUndo = true
Const blnRedo = true
Const blnLeftJustify = true
Const blnCentre = true
Const blnRightJustify = true
Const blnFullJustify = true
Const blnOrderList = true
Const blnUnOrderList = true
Const blnOutdent = true
Const blnIndent = true
Const blnAddHyperlink = true
Const blnAddImage = true
Const blnInsertTable = true
Const blnEmoticonPopUp = false
Const blnSpecialCharacters = true
Const blnPrint = true
Const blnStrikeThrough = true
Const blnSubscript = true
Const blnSuperscript = true
Const blnHorizontalRule = true
Const blnPreview = true
'***************************************************************************


'Advanced controls
'***************************************************************************
Const blnAdvAdddHyperlink = true 'Advanced hyperlink control
Const blnAdvAddImage = true 	'Advanced image control requires File System Object (FSO)
Const blnHTMLView = true	'Allows the user to view the HTML code, you may need to dsiable this for extra security
Const blnSpellCheck = true	'Requires IEspell for Ineternet Explorer or SpellBound for Mozilla
Const blnUseCSS = true		'Enable CSS (Cascading Style Sheets) in Mozilla
Const blnNoIEdblLine = true	'Prevent IE's standard double line spacing when the 'ENTER' key is pressed
'***************************************************************************



'Open and save configuration
'***************************************************************************
'For security reasons you shouldn't allow users to open/save server side scripting files with .asp, .php, etc. extensions as these could be used to execute malicious code on the server
'Saved files, for security, are only saved in the root folder you set below and not sub folders

Const blnOpenFile = true 'Open server file in editor (ie. open and HTML page in editor), requires File System Object (FSO)
Const strOpenFileFolderPath = "my_documents"
Const strOpenFileTypes = "htm;html" 'Separate file types with ;

Const blnSave = true 'Save editor contents to server, requires File System Object (FSO)
Const strSaveFileFolderPath = "my_documents\my_files"
Const strSaveFileTypes = "htm;html" 'Separate file types with ;
Const blnOverWriteFiles = false 'Set this to true only if you wish the user to overwrite saved files
'***************************************************************************



'Upload component (uncomment the upload component you require)
'***************************************************************************
'Select an upload component from the list below if you wish to use image upload and file attachement features
'The upoload component you select MUST be installed on the web server

Const strUploadComponent = "AspUpload"	'Persits AspUpload 3.x or above (includes progess bar) - http://www.aspupload.com
'Const strUploadComponent = "AspUpload2"	'Persits AspUpload - http://www.aspupload.com
'Const strUploadComponent = "Dundas"	'Dundas Upload - http://www.dundas.com 
'Const strUploadComponent = "fileUp"	'SA FileUp - http://www.softartisans.com
'Const strUploadComponent = "aspSmart"	'aspSmartUpload - http://www.aspsmart.com
'Const strUploadComponent = "AspSimple"	'AspSimpleUpload - http://www.asphelp.com
'***************************************************************************

'Upload component configuration
'***************************************************************************
'For security reasons you shouldn't allow users to upload files that could be used to execute malicious code on the server
'Uploaded files and images, for security, are only saved in the root folder you set below and not sub folders

'File Attachments
Const blnAttachments = false 'Attach files to the editor contents
Const strFileUploadPath = "my_documents\my_files"
Const strFileTypes = "zip;rar;doc;pdf;txt;rtf;htm;html;gif;jpg;png" 'Separate file types with ;
Const intMaxFileSize = 15360 'Use AspUpload 3.x or above for uploads over 2048KB (size in KB)

'Image uploads
Const blnImageUpload = true 'Upload images to the editor
'blnAdvAddImage (advanced image insert popup), also needs to be enabled to upload images as it uses the same pop up window 
Const strImageUploadPath = "../urunler"
Const strImageTypes = "jpeg;jpg;gif;png" 'Separate file types with ;
Const intMaxImageSize = 1024 'Use AspUpload 3.x or above for uploads over 2048KB (size in KB)
'***************************************************************************



'Using full URL path for images and links
'***************************************************************************
'If you are submitting the RTE content to a file outside of the RTE folder you may find that some of the relative
'paths for things like images stored on the server are incorrect (ie. href="my_documents/myPicture.jpg")
'The following can be used to change those relative server paths to full URL's so that if the submitted content is 
'displayed on a page out side of the RTE files the paths to images etc. still work

Const blnUseFullURLpath = false
Const strFullURLpathToRTEfiles = "" 'Type in the full URL to the RTE folder eg. "http://www.myweb.com/RTE/"

'***************************************************************************


'Basic image and CSS file location
'***************************************************************************
'Path to default images
Const strRTEImagePath = "images/"
Const strRTECSSfile = "css_styles/"
'***************************************************************************
%>