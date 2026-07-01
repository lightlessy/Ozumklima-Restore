<%
'****************************************************************************************
'**  Copyright Notice    
'**
'**  Web Wiz Rich Text Editor(TM)
'**  http://www.richtexteditor.org
'**                                               
'**  Copyright (C)2001-2012 Web Wiz(TM). All Rights Reserved. 
'**  
'**  THIS SOFTWARE IS PROVIDED BY THE COPYRIGHT HOLDERS UNDER LICENSE FROM 'WEB WIZ'.
'**  
'**  IF YOU DO NOT AGREE TO THE LICENSE AGREEMENT THEN 'WEB WIZ' IS UNWILLING TO LICENSE 
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
'**  Web Wiz, Unit 10E, Dawkins Road Industrial Estate, Poole, Dorset, BH15 4JD, England
'**  http://www.webwiz.co.uk
'**
'**  Removal or modification of this copyright notice will violate the license contract.
'**
'****************************************************************************************



'Uncomment the character set you require 
'(uncomment means remove the (') single quote from infront of the line)

'Western European ISO
'Const strPageEncoding = "iso-8859-1"

'Unicode UTF-8
Const strPageEncoding = "utf-8"

'Arabic ISO
'Const strPageEncoding = "iso-8859-6"

'Arabic Windows
'Const strPageEncoding = "windows-1256"

'Baltic Windows
'Const strPageEncoding = "windows-1257"

'Central European DOS
'Const strPageEncoding = "ibm852"

'Central European ISO
'Const strPageEncoding = "iso-8859-2"

'Central European Windows
'Const strPageEncoding = "windows-1250"

'Chinese Simplified
'Const strPageEncoding = "gb2312"

'Chinese Simplified
'Const strPageEncoding = "hz-gb-2312"

'Chinese Traditional
'Const strPageEncoding = "big5"

'Cyrillic ISO
'Const strPageEncoding = "iso-8859-5"

'Cyrillic KOI8-R
'Const strPageEncoding = "koi8-r"

'Cyrillic KOI8-U
'Const strPageEncoding = "koi8-ru"

'Cyrillic Windows
'Const strPageEncoding = "windows-1251"

'Greek ISO
'Const strPageEncoding = "iso-8859-7"

'Greek Windows
'Const strPageEncoding = "windows-1253"

'Hebrew ISO-Logical
'Const strPageEncoding = "iso-8859-8-i"

'Hebrew ISO-Visual
'Const strPageEncoding = "iso-8859-8"

'Hebrew Windows
'Const strPageEncoding = "windows-1255"

'Japanese EUC
'Const strPageEncoding = "euc-jp"

'Japanese Shift-JIS
'Const strPageEncoding = "shift-jis"

'Korean
'Const strPageEncoding = "euc-kr"

'Thai Windows
'Const strPageEncoding = "windows-874"

'Turkish ISO
'Const strPageEncoding = "iso-8859-9"

'Vietnamese
'Const strPageEncoding = "windows-1258"

'Write the page encoding meta tage
Response.Write("<meta http-equiv=""Content-Type"" content=""text/html; charset=" & strPageEncoding & """>")

%>