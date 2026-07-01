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





'********************************************
'***  Rich Text Compatible Browser type *****
'********************************************

Private Function RTEenabled()

	Dim strUserAgent	'Holds info on the users browser
	Dim strMSIEversion	'Holds the version of IE

	'Get the users HTTP user agent (web browser)
	strUserAgent = Request.ServerVariables("HTTP_USER_AGENT")
	
	'Response.Write(strUserAgent)
	

	'*************************************
	'***** Windows Internet Explorer *****
	'*************************************
	
	'See if the user agent is IE on Winows and not Opera trying to look like IE
	If InStr(1, strUserAgent, "MSIE", 1) > 0 AND InStr(1, strUserAgent, "Win", 1) > 0 AND InStr(1, strUserAgent, "Opera", 1) = 0 Then
		
		'Get the MSIE version
		strMSIEversion = Trim(Mid(strUserAgent, inStr(1, strUserAgent, "MSIE", 1)+5, 3))
		
		'Now we know this is Windows IE we need to see if the version number is 5.5
		If strMSIEversion = "5.5" OR strMSIEversion = "5,5" Then

			RTEenabled = "winIE"
		
		'Now we know this is Windows IE we need to see if the version number is 6+ (error handling to make sure number is numeric)
		ElseIf isNumeric(strMSIEversion) Then
			
			'Now check the version number is 6 or above
			If CInt(strMSIEversion) >= 6 Then
				RTEenabled = "winIE"
			'Else IE is below 5
			Else
				RTEenabled = "false"
			End If

		'Else the IE version is below 5 so return na
		Else

			RTEenabled = "false"
		End If
		
		
	'****************************
	'***** Mozilla Firebird *****
	'****************************
	
	'See if this is a version of Mozilla Firebird that supports Rich Text Editing under it's Midas API
	ElseIf inStr(1, strUserAgent, "Firebird", 1) Then
		
		'Now we know this is Mozilla Firebird we need to see if the version 0.6.1 or above; relase date is above 2003/07/28
		If CLng(Trim(Mid(strUserAgent, CInt(inStr(1, strUserAgent, "Gecko/", 1)+6), 8))) >= 20030728 Then

			RTEenabled = "Gecko"

		'Else the Mozilla Firebird version is below 1.5 so return false
		Else

			RTEenabled = "false"
		End If
	
	
	'**********************************************
	'***** 	    Mozilla Firefox/Netscape      *****
	'**********************************************
	
	'See if this is a version of Mozilla/Netscape that supports Rich Text Editing under it's Midas API
	ElseIf inStr(1, strUserAgent, "Gecko", 1) > 0 AND inStr(1, strUserAgent, "Firebird", 1) = 0 AND isNumeric(Trim(Mid(strUserAgent, CInt(inStr(1, strUserAgent, "Gecko/", 1)+6), 8))) Then
		
		'Now we know this is Mozilla/Netscape we need to see if the version number is above 1.3 or above; relase date is above 2003/03/12
		If CLng(Trim(Mid(strUserAgent, CInt(inStr(1, strUserAgent, "Gecko/", 1)+6), 8))) => 20030312 Then

			RTEenabled = "Gecko"

		'Else the Mozilla version is below 1.3 or below 7.1 of Netscape so return false
		Else

			RTEenabled = "false"
		End If
	
	
	
	'**********************************************
	'***** 		Opera			  *****
	'**********************************************
	
	'See if this is Opera that supports Rich Text (Opera 9 and above)
	ElseIf inStr(1, strUserAgent, "Opera", 1) AND InStr(1, strUserAgent, "Opera Mini", 1) = 0 Then
		
		'now we need to see what version of Opera we are using
		If isNumeric(Trim(Mid(strUserAgent, CInt(inStr(1, strUserAgent, "Opera/", 1)+6), 1))) Then
			
			If CLng(Trim(Mid(strUserAgent, CInt(inStr(1, strUserAgent, "Opera/", 1)+6), 1))) => 9 Then
				
				RTEenabled = "opera"
			
			'Else the Opera version is below 9 so return false
			Else
	
				RTEenabled = "false"
			End If
		
		'Else the Opera version is below 9 so return false
		Else
	
			RTEenabled = "false"
		End If
		
	
	'******************************
	'***** 	  Google Chrome	  *****
	'******************************
	
	'See if this is Chrome
	ElseIf inStr(1, strUserAgent, "Chrome", 1) Then
		
		'RTE is not supported on Android
		If inStr(1, strUserAgent, "Android", 1) Then  
			
			RTEenabled = "false"
		
		'Else all other versions of Chrome do support RTE
		Else
			
			'Chrome works idenetically to Gekco so idenetify as Gecko
			RTEenabled = "Gecko"
			
	
		End If
	
	
	'**********************************************
	'***** 	  Apple Safari & Google Chrome	  *****
	'**********************************************
	
	'See if this is The AppleWebKit Engine that supports Rich Text (Safari 3.1 and above or Google Chrome)
	ElseIf inStr(1, strUserAgent, "AppleWebKit", 1) Then
		
		'Javascript is not supported on the iPhone, iPod, iPad, and Android
		If inStr(1, strUserAgent, "iPhone", 1) OR inStr(1, strUserAgent, "iPad", 1) OR inStr(1, strUserAgent, "iPod", 1) OR inStr(1, strUserAgent, "Android", 1) Then  
			
			RTEenabled = "false"
		
		'See what everion we are using of AppleWebKit
		ElseIf isNumeric(Trim(Mid(strUserAgent, CInt(inStr(1, strUserAgent, "AppleWebKit/", 1)+12), 3))) Then
			
			'If the version number is 523 or above it is fully RTE enabled
			If CLng(Trim(Mid(strUserAgent, CInt(inStr(1, strUserAgent, "AppleWebKit/", 1)+12), 3))) => 523 Then
				
				'AppleWebKit engine works idenetically to the Gekco engine so idenetify as Gecko
				RTEenabled = "Gecko"
			
			'Else the the version is older and not fully RTE enabled
			Else
	
				RTEenabled = "false"
			End If
		
		'Else the Safari version is below 3.0.4 so return false
		Else
	
			RTEenabled = "false"
		End If
	
	
	
	'***********************************
	'***** Non RTE Enabled Browser *****
	'***********************************
	
	'Else this is a browser that does not support Rich Text Editing
	Else
		'RTEenabled - false
		RTEenabled = "false"
	End If
	'RTEenabled = "false"
	
End Function





'******************************************
'***  	   Random Hex Generator        ****
'******************************************

Private Function hexValue(ByVal intHexLength)

	Dim intLoopCounter
	Dim strHexValue

	'Randomise the system timer
	Randomize Timer()

	'Generate a hex value
	For intLoopCounter = 1 to intHexLength

		'Genreate a radom decimal value form 0 to 15
		intHexLength = CInt(Rnd * 1000) Mod 16

		'Turn the number into a hex value
		Select Case intHexLength
			Case 1
				strHexValue = "1"
			Case 2
				strHexValue = "2"
			Case 3
				strHexValue = "3"
			Case 4
				strHexValue = "4"
			Case 5
				strHexValue = "5"
			Case 6
				strHexValue = "6"
			Case 7
				strHexValue = "7"
			Case 8
				strHexValue = "8"
			Case 9
				strHexValue = "9"
			Case 10
				strHexValue = "A"
			Case 11
				strHexValue = "B"
			Case 12
				strHexValue = "C"
			Case 13
				strHexValue = "D"
			Case 14
				strHexValue = "E"
			Case 15
				strHexValue = "F"
			Case Else
				strHexValue = "Z"
		End Select

		'Place the hex value into the return string
		hexValue = hexValue & strHexValue
	Next
End Function




'******************************************
'***  	   Unsafe character Strip      ****
'******************************************

'Function to strip non alphanumeric characters email addresses
Private Function characterStrip(ByVal strTextInput)

	'Dimension variable
	Dim intLoopCounter 	'Holds the loop counter
	
	'Loop through the ASCII characters
	For intLoopCounter = 0 to 31
		strTextInput = Replace(strTextInput, CHR(intLoopCounter), "", 1, -1, 0) 
	Next
	
	'Loop through the ASCII characters
	For intLoopCounter = 33 to 37
		strTextInput = Replace(strTextInput, CHR(intLoopCounter), "", 1, -1, 0) 
	Next
	
	'Loop through the ASCII characters
	For intLoopCounter = 39 to 44
		strTextInput = Replace(strTextInput, CHR(intLoopCounter), "", 1, -1, 0) 
	Next
	
	'Loop through the ASCII characters
	For intLoopCounter = 58 to 65
		strTextInput = Replace(strTextInput, CHR(intLoopCounter), "", 1, -1, 0) 
	Next
	
	'Loop through the ASCII characters numeric characters
	For intLoopCounter = 91 to 94
		strTextInput = Replace(strTextInput, CHR(intLoopCounter), "", 1, -1, 0) 
	Next
	
	'Loop through the extended ASCII characters
	For intLoopCounter = 123 to 125
		strTextInput = Replace(strTextInput, CHR(intLoopCounter), "", 1, -1, 0) 
	Next
	
	'Loop through the extended ASCII characters
	For intLoopCounter = 127 to 255
		strTextInput = Replace(strTextInput, CHR(intLoopCounter), "", 1, -1, 0) 
	Next
	
	'Strip individul ASCII characters left out from above
	strTextInput = Replace(strTextInput, CHR(59), "", 1, -1, 0)
	strTextInput = Replace(strTextInput, CHR(60), "", 1, -1, 0)
	strTextInput = Replace(strTextInput, CHR(62), "", 1, -1, 0)
	strTextInput = Replace(strTextInput, CHR(96), "", 1, -1, 0)
	
	
	'Return the string
	characterStrip = strTextInput
	
End Function




'*********************************************
'***  		Strip all tags		 *****
'*********************************************

'Remove all tags for text only display 
Private Function removeAllTags(ByVal strInputEntry)

	'Remove all HTML scripting tags etc. for plain text output
	strInputEntry = Replace(strInputEntry, ";", "", 1, -1, 1)
	strInputEntry = Replace(strInputEntry, "&", "&amp;", 1, -1, 1)
	strInputEntry = Replace(strInputEntry, "<", "&lt;", 1, -1, 1)
	strInputEntry = Replace(strInputEntry, ">", "&gt;", 1, -1, 1)
	strInputEntry = Replace(strInputEntry, "'", "&#039;", 1, -1, 1)
	strInputEntry = Replace(strInputEntry, """", "&quot;", 1, -1, 1)
	strInputEntry = Replace(strInputEntry, "\", "&#092;", 1, -1, 1)

	'Return
	removeAllTags = strInputEntry
End Function






'Get about status
Private Function rteAbout()

	rteAbout = 1 

	'Calulate the lentgh
	If LEN(strLicense) > 47 Then
		
		Dim intAbout
		intAbout = 1
		
		'Read in the data
		If isNumeric(Mid(strLicense, len(strLicense)-3, 4)) Then intAbout = CInt(Mid(strLicense, len(strLicense)-3, 4))
		strDisplayLicense = Mid(strLicense, 4, 40)
		rteAbout = CBool(intAbout MOD 24) 
	End If
End Function
 

'Get RTE info
Private Sub rteInfo()

	Response.Write("" & _
	vbCrLf & "<pre style=""text-align:left;"">" & _
	vbCrLf & "*********************************************************" & _
	vbCrLf & "Software: Web Wiz Rich Text Editor" & _
	vbCrLf & "Version: " & strRTEversion & _
	vbCrLf & "License: " & strDisplayLicense & _
	vbCrLf & "Author: Web Wiz." & _
	vbCrLf & "Address: Unit 10E, Dawkins Raod Ind Est, Poole, Dorset, UK" & _
	vbCrLf & "Info: http://www.richtexteditor.org" & _
	vbCrLf & "Copyright: (C)2001-2012 Web Wiz Ltd. All rights reserved" & _
	vbCrLf & "*********************************************************" & _
	vbCrLf & "</pre")
	
	Response.Flush
	Response.End
End Sub



'******************************************
'***  	Convertion Functions	 	***
'******************************************

'CInt Handling Integers to 32,768
Private Function IntC(strExpression) 

	'Set error trapping
	On Error Resume Next
	
	'Converts the string data to an Integer Number
	IntC = CInt(strExpression)
	
	'If an error has occurred write an error to the page
	If Err.Number <> 0 Then
		
		Response.Write("RTE Number handling error; The data being converted is not within range; -32,768 to 32,768.")
		Response.Flush
		Response.End
	End If 
	
	'Disable error trapping
	On Error goto 0

End Function


'CLng Handling Integers to 2,147,483,648
Private Function LngC(strExpression) 

	'Set error trapping
	On Error Resume Next
	
	'Converts the string data to an Integer Number
	LngC = CLng(strExpression)
	
	'If an error has occurred write an error to the page and stop
	If Err.Number <> 0 Then	
		
		Response.Write("RTE Number handling error; The data being converted is not within range; -2,147,483,648 to 2,147,483,648.")
		Response.Flush
		Response.End
	End If
	
	'Disable error trapping
	On Error goto 0

End Function


'CDbl Handling Floating Point Numbers
Private Function DblC(strExpression) 

	'Set error trapping
	On Error Resume Next
	
	'Converts the string data to an Integer Number
	DblC = CDbl(strExpression)
	
	'If an error has occurred write an error to the page
	If Err.Number <> 0 Then	
		
		Response.Write("RTE Number handling error; The data being converted is not a valid Floating Point Number.")
		Response.Flush
		Response.End
	End If
	
	'Disable error trapping
	On Error goto 0

End Function


'CBool Handling True and False
Private Function BoolC(strExpression) 

	'Set error trapping
	On Error Resume Next
	
	'Converts the string data to an Integer Number
	BoolC = CBool(strExpression)
	
	'If an error has occurred write an error to the page
	If Err.Number <> 0 Then	
		
		Response.Write("RTE Expression handling error; The data being converted is not a valid Boolean Subtype.")
		Response.Flush
		Response.End
	End If
	
	'Disable error trapping
	On Error goto 0

End Function


'CDate Handling Date Subtypes
Private Function DateC(strExpression) 

	'Set error trapping
	On Error Resume Next
	
	'Converts the string data to an Integer Number
	DateC = CDate(strExpression)
	
	'If an error has occurred write an error to the page
	If Err.Number <> 0 Then
		
		Response.Write("RTE Expression handling error; The data being converted is not a valid Date.")
		Response.Flush
		Response.End
	End If
		
	'Disable error trapping
	On Error goto 0

End Function
%>