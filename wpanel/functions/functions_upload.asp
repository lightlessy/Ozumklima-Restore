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


'******************************************
'***	   File Upload Function        ****
'******************************************

'Function to upload a file
Private Function fileUpload(ByVal strFileUploadPath, ByVal saryFileUploadTypes, ByVal intMaxFileSize, ByVal strUploadComponent, ByRef lngErrorFileSize, ByRef blnExtensionOK)

	'Dimension variables
	Dim objUpload		'Uplaod component
	Dim strNewFileName	'Holds the file name
	Dim strOriginalFileName	'Holds the original file name for those components that need to save the file first
	Dim objFSO		'Holds the file system object
	Dim file		'Holds the FSO file object
	


	'******************************************
	'***	     Upload components         ****
	'******************************************

	'Select which upload component to use
	Select Case strUploadComponent
	

		'******************************************
		'***     Persits AspUpload component   ****
		'******************************************

		'Persits AspUpload upload component - tested with version 3.0
		Case "AspUpload", "AspUpload2"

			'Create upload object
		
			Set objUpload = Server.CreateObject("Persits.Upload")   
			
			'If AspUpload 3.x or above get the progress ID for the progress bar
			If strUploadComponent = "AspUpload" Then objUpload.ProgressID = Request.QueryString("PID")

			With objUpload


				'make sure files arn't over written
				.OverwriteFiles = False

				'We need to save the file before we can find out anything about it
				'** Save virtual is used as saving to memory is often disabled by the web host **
				.SaveVirtual strFileUploadPath

				'Get the file name
				strNewFileName = .Files(1).ExtractFileName

				'Replace spaces with underscores
				strNewFileName = Replace(strNewFileName, " ", "_", 1, -1, 1)
				
				'Remove semicolons to prevent IIS 6 and below vulnerbility
				strNewFileName = Replace(strNewFileName, ";", "", 1, -1, 1)

				'Remove brackets that aspUpload may put in when renaming a file
				strNewFileName = Replace(strNewFileName, "(", "", 1, -1, 1)
				strNewFileName = Replace(strNewFileName, ")", "", 1, -1, 1)

				'Check the file size is not above the max allowed size, this is done using a function not the compoent to stop an exception error
				lngErrorFileSize = fileSize(.Files(1).Size, intMaxFileSize)

				'Loop through all the allowed extensions and see if the file has one
				blnExtensionOK = fileExtension(strNewFileName, saryFileUploadTypes)

				'If the file is OK save it to disk
				If lngErrorFileSize = 0 AND blnExtensionOK = True Then

					'** AspUpload does have a function to stop files form being over written but this can allow the rte open to malicious code **
					'** so instead if the file exists create an acceptable new name for the file **
					'Check to make sure the file does not already exist
					Do While .FileExists(Server.MapPath(strFileUploadPath) & "\" & strNewFileName)

						'Create a new file name for the file if it already exsist
						strNewFileName = hexValue(3) & "_" & strNewFileName
					Loop
					

					'Save the file to disk with new file name
					'** Copy virtual is used as save as is often disabled by the web host **
					.Files(1).CopyVirtual strFileUploadPath & "/" & strNewFileName

					'As a new copy of the file is saved we need to get rid of the old copy
					.Files(1).Delete

					'Pass the filename back
					fileUpload = strNewFileName


				'Else if it is not OK delete the uploaded file
				Else
					.Files(1).Delete

				End If

			End With

			'Clean up
			Set objUpload = Nothing




		'******************************************
		'***         Dundas Upload component   ****
		'******************************************

		'Dundas upload component free from http://www.dundas.com - tested with version 2.0
		Case "Dundas"

			'Create upload object
			Set objUpload = Server.CreateObject("Dundas.Upload")

			With objUpload

				'Make sure we are using a virtual directory for script
				.UseVirtualDir = True

				'Make sure the file names are not unique at this time
				.UseUniqueNames = False

				'Save the file first to memory
				.SaveToMemory()

				'Get the file name, the path mehod will be empty as we are saving to memory so use the original file path of the users system to get the name
				strNewFileName = .GetFileName(.Files(0).OriginalPath)

				'Replace spaces with underscores
				strNewFileName = Replace(strNewFileName, " ", "_", 1, -1, 1)
				
				'Remove semicolons to prevent IIS 6 and below vulnerbility
				strNewFileName = Replace(strNewFileName, ";", "", 1, -1, 1)

				'Check the file size is not above the max allowed size, this is done using a function not the compoent to stop an exception error
				lngErrorFileSize = fileSize(.Files(0).Size, intMaxFileSize)

				'Loop through all the allowed extensions and see if the file has one
				blnExtensionOK = fileExtension(strNewFileName, saryFileUploadTypes)

				'If the file is OK save it to disk
				If lngErrorFileSize = 0 AND blnExtensionOK = True Then

					'** DunDas Upload does have a function create unique file names but this can allow thf orum open to malicious code **
					'** so instead if the file exists create an acceptable new name for the file **
					'Check to make sure the file does not already exist
					Do While .FileExists(strFileUploadPath & "\" & strNewFileName)

						'Create a new file name for file if it already exsist
						strNewFileName = hexValue(3) & "_" & strNewFileName
					Loop
					
					
					'Save the file to disk
					.Files(0).SaveAs strFileUploadPath & "/" & strNewFileName

					'Pass the filename back
					fileUpload = strNewFileName
				End If
			End With

			'Clean up
			Set objUpload = Nothing




		'******************************************
		'***  SoftArtisans FileUp component    ****
		'******************************************

		'SA FileUp upload component - tested with version 4
		Case "fileUp"

			'Dimension variables
			Dim objSAFileManager	'Holds the file mangare object to check to see if the file already exists

			'Create upload object
			Set objUpload = Server.CreateObject("SoftArtisans.FileUp")

			With objUpload

				'Over write files or an exception will occur if it already exists
				.OverWriteFiles = True

				'Set the upload path
				.Path = Server.MapPath(strFileUploadPath)

				'Get the file name, the path mehod will be empty as we are saving to memory so use the original file path of the users system to get the name
				strNewFileName = Mid(.UserFilename, InstrRev(.UserFilename, "\") + 1)

				'Replace spaces with underscores
				strNewFileName = Replace(strNewFileName, " ", "_", 1, -1, 1)
				
				'Remove semicolons to prevent IIS 6 and below vulnerbility
				strNewFileName = Replace(strNewFileName, ";", "", 1, -1, 1)

				'Check the file size is not above the max allowed size, this is done using a function not the compoent to stop an exception error
				lngErrorFileSize = fileSize(.TotalBytes, intMaxFileSize)

				'Loop through all the allowed extensions and see if the file has one
				blnExtensionOK = fileExtension(strNewFileName, saryFileUploadTypes)

				'If the file is OK save it to disk
				If lngErrorFileSize = 0 AND blnExtensionOK = True Then

					'Create a file manager object to get file access to check if the file already exists
					Set objSAFileManager = Server.CreateObject("SoftArtisans.FileManager")

					'Check to make sure the file does not already exist
					Do While objSAFileManager.FileExists(Server.MapPath(strFileUploadPath) & "\" & strNewFileName)

						'Create a new file name for file if it already exsist
						strNewFileName = hexValue(3) & "_" & strNewFileName
					Loop
					
					
					'Clean up
					Set objSAFileManager = Nothing

					'Save the file to disk
					.SaveAs strNewFileName

					'Pass the filename back
					fileUpload = strNewFileName
				End If

			End With

			'Clean up
			Set objUpload = Nothing




		'******************************************
		'***  	AspSmartUpload component       ****
		'******************************************

		'AspSmartUpload upload component free from http://www.aspsmart.com
		Case "aspSmart"

			'Create upload object
			Set objUpload = Server.CreateObject("aspSmartUpload.SmartUpload")

			With objUpload

				'Make sure we are using a virtual directory for script
				.DenyPhysicalPath = True

				'Save the file first to memory
				.Upload

				'Get the file name, the path mehod will be empty as we are saving to memory so use the original file path of the users system to get the name
				strNewFileName = .Files(1).Filename

				'Replace spaces with underscores
				strNewFileName = Replace(strNewFileName, " ", "_", 1, -1, 1)
				
				'Remove semicolons to prevent IIS 6 and below vulnerbility
				strNewFileName = Replace(strNewFileName, ";", "", 1, -1, 1)

				'Check the file size is not above the max allowed size
				lngErrorFileSize = fileSize(.Files(1).Size, intMaxFileSize)

				'Loop through all the allowed extensions and see if the file has one
				blnExtensionOK = fileExtension(strNewFileName, saryFileUploadTypes)

				'If the file is OK save it to disk
				If lngErrorFileSize = 0 AND blnExtensionOK = True Then

					'Create a new file name for file as using a random genrated hex code, and hope it doesn't already exist as there is noway with this component of finding if the file already exists
					strNewFileName = hexValue(5) & "_" & strNewFileName
					
					
					'Save the file to disk
					.Files(1).SaveAs strFileUploadPath & "/" & strNewFileName

					'Pass the filename back
					fileUpload = strNewFileName
				End If

			End With

			'Clean up
			Set objUpload = Nothing



		'******************************************
		'***     AspSimpleUpload component     ****
		'******************************************

		'ASPSimpleUpload component
		Case "AspSimple"


			'Create upload object
			Set objUpload = Server.CreateObject("ASPSimpleUpload.Upload")

			With objUpload

				'Get the file name
				strOriginalFileName = .ExtractFileName(.Form("file"))

				'Save the amended file name
				strNewFileName = "TMP" & hexValue(7) & "_" & strOriginalFileName

				'Save the file to disk first so we can check it
				Call .SaveToWeb ("file", strFileUploadPath & "/" & strNewFileName)


				'Create the file system object
				Set objFSO = Server.CreateObject("Scripting.FileSystemObject")

				'Create a file object with the file details
				Set file = objFSO.GetFile(Server.MapPath(strFileUploadPath) & "\" & strNewFileName)

				'Check the file size is not above the max allowed size, this is done using a function not the compoent to stop an exception error
				lngErrorFileSize = fileSize(file.Size, intMaxFileSize)


				'Place the original file name back in the new filename variable
				strNewFileName = strOriginalFileName

				'Replace spaces with underscores
				strNewFileName = Replace(strNewFileName, " ", "_", 1, -1, 1)
				
				'Remove semicolons to prevent IIS 6 and below vulnerbility
				strNewFileName = Replace(strNewFileName, ";", "", 1, -1, 1)


				'Loop through all the allowed extensions and see if the file has one
				blnExtensionOK = fileExtension(strNewFileName, saryFileUploadTypes)


				'If the file is OK save it to disk
				If lngErrorFileSize = 0 AND blnExtensionOK = True Then

					'Check to make sure the file does not already exist
					Do While objFSO.FileExists(Server.MapPath(strFileUploadPath) & "\" & strNewFileName)

						'Create a new file name for file if it already exsist
						strNewFileName = hexValue(3) & "_" & strNewFileName
					Loop
					
					'Save the file to disk
					Call .SaveToWeb("file", strFileUploadPath & "/" & strNewFileName)

					'Delete the original file
					file.Delete

					'Pass the filename back
					fileUpload = strNewFileName
				End If

			End With

			'Clean up
			Set file = Nothing
			Set objFSO = Nothing
			Set objUpload = Nothing

	End Select

End Function





'******************************************
'***	Check file size function       ****
'******************************************
Function fileSize(ByVal lngFileSize, ByVal intMaxFileSize)

	'If the file size is to large place the present file size in then return the file size
	If CLng(lngFileSize / 1024) > intMaxFileSize Then

		fileSize = CLng(lngFileSize / 1024)

	'Else set the return value to 0
	Else
		fileSize = 0
	End If

End Function





'******************************************
'***	Check file ext. function       ****
'******************************************
Function fileExtension(ByVal strFileName, ByVal saryFileUploadTypes)

	'Dimension varibles
	Dim intExtensionLoopCounter

	'Intilaise return value
	fileExtension = False

	'Loop through all the allowed extensions and see if the file has one
	For intExtensionLoopCounter = 0 To UBound(saryFileUploadTypes)

		If LCase(Right(strFileName, Len(saryFileUploadTypes(intExtensionLoopCounter))+1)) = "." & LCase(saryFileUploadTypes(intExtensionLoopCounter)) Then fileExtension = True
	Next

End Function
%>