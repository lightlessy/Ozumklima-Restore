<%
	Dim StrDate
	Dim StrTime
	Dim StrDate_Time
	Dim StrWeekDayName
	Dim StrMonthName
	
	On Error Resume Next

	StrDate = FormatDateTime(Date(),vbshortdate)
	StrTime = FormatDateTime(Now(),vbshorttime)
	StrDate_Time = Now()
	StrWeekDayName = WeekdayName(weekday(date))
	StrMonthName = MonthName(month(date))

	Set ObjSettings = Createobject("ADODB.Recordset")
		ObjSettings.open  "SELECT * FROM TblSettings", ObjConn , 3 , 1

		StrTitle = ObjSettings("StrTitle")
		StrVer = ObjSettings("StrVer")
		StrURL = ObjSettings("StrURL")
		StrEmail = ObjSettings("StrEmail")
		StrLogo = ObjSettings("StrLogo")
		StrAdminUserName = ObjSettings("StrAdminUserName")
		StrAdminUserPass = ObjSettings("StrAdminUserPass")
		StrUSDCur = ObjSettings("StrUSDCur")
		StrEUROCur = ObjSettings("StrEUROCur")
		StrProdList = ObjSettings("StrProdList")
		StrFooter = ObjSettings("StrFooter")
		StrBannerURL = ObjSettings("StrBannerURL")
		StrBannerPath = ObjSettings("StrBannerPath")
		
		StrDomain = "http://www.cemexnaturalstone.com"

		ObjSettings.Close()
	Set ObjSettings = Nothing

	'##---

	If Request.Cookies(StrTitle)("StrFirstVisit") = "" Then
		Response.Cookies(StrTitle)("StrFirstVisit") = Now()
		Response.Cookies(StrTitle).Expires = Now() + 365  
	Else
		Response.Cookies(StrTitle)("StrLastVisit") = Now()
		Response.Cookies(StrTitle).Expires = Now() + 365 
	End If

	'##---

	Function StrRandomCode(StrNumber)
		
		For i = 1 To (StrNumber/2) 
			Randomize
			StrCode = StrCode + Chr(Int((10) * Rnd + 48))
			StrCode = StrCode + Chr(Int((26) * Rnd + 65))
			
			StrRandomCode = StrCode
		Next

	End Function

	'##---

	Sub StrDbModified

		Dim StrModified

		Set ObjFSO = Server.CreateObject("Scripting.FileSystemObject")
		Set ObjPath = ObjFSO.GetFile(Server.MapPath("db/database.mdb"))
			StrModified = ObjPath.DateLastModified

			Response.Write "Son güncelleme : " & StrModified

		Set ObjFSO = Nothing
		Set ObjPath = Nothing

	End Sub

	'##---


%>