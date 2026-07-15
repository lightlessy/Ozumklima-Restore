<%@ Language="VBScript" CodePage="65001" %>
<%
Response.Buffer = True
Response.CodePage = 65001
Response.Charset = "utf-8"

If Request.Cookies("ozum")("wpAfLogin") <> "true" Then
    Response.Redirect "index.asp"
    Response.End
End If
%>

<!--#include file="dbase.asp"-->

<%
Function SqlText(ByVal value)
    SqlText = Replace(CStr(value & ""), "'", "''")
End Function

Function FormValue(ByVal fieldName)
    FormValue = Trim(Request.Form(fieldName) & "")
End Function

Function IntegerOrDefault(ByVal value, ByVal defaultValue)
    If IsNumeric(value) Then
        IntegerOrDefault = CLng(value)
    Else
        IntegerOrDefault = CLng(defaultValue)
    End If
End Function

Sub StopWithError(ByVal message, ByVal groupId, ByVal typeId)
    Dim returnUrl
    returnUrl = "join.asp"

    If IsNumeric(groupId) Then
        returnUrl = returnUrl & "?g=" & CLng(groupId)

        If IsNumeric(typeId) Then
            returnUrl = returnUrl & "&tip=" & CLng(typeId)
        End If
    End If

    Response.Status = "400 Bad Request"
    Response.Write "<!doctype html><html><head><meta charset=""utf-8""><title>Ürün kaydedilemedi</title>"
    Response.Write "<style>body{font-family:Arial,sans-serif;background:#f4f4f4;color:#333;padding:40px}.box{max-width:720px;margin:auto;background:#fff;border:1px solid #ddd;padding:28px}.error{color:#8b2424}.button{display:inline-block;margin-top:18px;padding:10px 16px;background:#3c3c3c;color:#fff;text-decoration:none}</style>"
    Response.Write "</head><body><div class=""box""><h2 class=""error"">Ürün kaydedilemedi</h2>"
    Response.Write "<p>" & Server.HTMLEncode(message) & "</p>"
    Response.Write "<a class=""button"" href=""" & Server.HTMLEncode(returnUrl) & """>Ürün formuna dön</a>"
    Response.Write "</div></body></html>"
    Response.End
End Sub

If UCase(Request.ServerVariables("REQUEST_METHOD")) <> "POST" Then
    Response.Redirect "join.asp"
    Response.End
End If

Dim isim, tip, anagrup, surface, notlar, delivery, yayin
Dim kodu, stok, keyw, descr, renk1, renk2, renk3, renk4, renk5
Dim sql1, checkRs, identityRs, yeniID, errorDescription

isim = FormValue("isim")
tip = FormValue("tip")
anagrup = FormValue("anagrup")
surface = IntegerOrDefault(FormValue("surface"), 0)
notlar = Request.Form("notlar") & ""
delivery = IntegerOrDefault(FormValue("delivery"), 3)
yayin = IntegerOrDefault(FormValue("yayin"), 0)
kodu = FormValue("kodu")
keyw = Request.Form("keyw") & ""
descr = Request.Form("descr") & ""
renk1 = FormValue("renk1")
renk2 = FormValue("renk2")
renk3 = FormValue("renk3")
renk4 = FormValue("renk4")
renk5 = FormValue("renk5")
stok = 10

If kodu = "" Then kodu = "0"
If renk1 = "" Then renk1 = "0"
If renk2 = "" Then renk2 = "0"
If renk3 = "" Then renk3 = "0"
If renk4 = "" Then renk4 = "0"
If renk5 = "" Then renk5 = "0"
If delivery < 0 Then delivery = 3
If yayin <> 1 Then yayin = 0

If isim = "" Then
    StopWithError "Ürün adı boş bırakılamaz.", anagrup, tip
End If

If Not IsNumeric(anagrup) Then
    StopWithError "Geçerli bir ana kategori seçilmedi.", anagrup, tip
End If

If Not IsNumeric(tip) Then
    StopWithError "Geçerli bir alt kategori seçilmedi.", anagrup, tip
End If

anagrup = CLng(anagrup)
tip = CLng(tip)

Set checkRs = baglanti.Execute( _
    "SELECT COUNT(*) AS toplam FROM tip " & _
    "WHERE id=" & tip & " AND grup=" & anagrup _
)

If checkRs("toplam") = 0 Then
    checkRs.Close
    Set checkRs = Nothing
    StopWithError "Seçilen alt kategori bu ana kategoriye bağlı değil.", anagrup, tip
End If

checkRs.Close
Set checkRs = Nothing

If surface > 0 Then
    Set checkRs = baglanti.Execute( _
        "SELECT COUNT(*) AS toplam FROM surface " & _
        "WHERE id=" & surface & " AND grup=" & anagrup _
    )

    If checkRs("toplam") = 0 Then
        checkRs.Close
        Set checkRs = Nothing
        StopWithError "Seçilen ürün grubu bu ana kategoriye bağlı değil.", anagrup, tip
    End If

    checkRs.Close
    Set checkRs = Nothing
End If

sql1 = "INSERT INTO products " & _
       "(isim, tip, anagrup, surface, renk1, renk2, renk3, renk4, renk5, " & _
       "notlar, delivery, yayin, kodu, stok, keyw, descr) " & _
       "VALUES (" & _
       "'" & SqlText(isim) & "', " & _
       tip & ", " & _
       anagrup & ", " & _
       surface & ", " & _
       "'" & SqlText(renk1) & "', " & _
       "'" & SqlText(renk2) & "', " & _
       "'" & SqlText(renk3) & "', " & _
       "'" & SqlText(renk4) & "', " & _
       "'" & SqlText(renk5) & "', " & _
       "'" & SqlText(notlar) & "', " & _
       delivery & ", " & _
       yayin & ", " & _
       "'" & SqlText(kodu) & "', " & _
       stok & ", " & _
       "'" & SqlText(keyw) & "', " & _
       "'" & SqlText(descr) & "'" & _
       ")"

On Error Resume Next
baglanti.BeginTrans
baglanti.Execute sql1

If Err.Number <> 0 Then
    errorDescription = Err.Description
    Err.Clear
    baglanti.RollbackTrans
    On Error GoTo 0
    StopWithError "Veritabanı kaydı sırasında hata oluştu: " & errorDescription, anagrup, tip
End If

Set identityRs = Nothing
Err.Clear
Set identityRs = baglanti.Execute("SELECT @@IDENTITY AS YeniID")

If Err.Number <> 0 Or identityRs Is Nothing Then
    Err.Clear
    Set identityRs = Nothing
    Set identityRs = baglanti.Execute("SELECT TOP 1 AffiliateID AS YeniID FROM products ORDER BY AffiliateID DESC")
End If

If Err.Number <> 0 Or identityRs Is Nothing Then
    errorDescription = Err.Description
    Err.Clear
    baglanti.RollbackTrans
    On Error GoTo 0
    StopWithError "Ürün kaydedildi ancak yeni ürün numarası alınamadı: " & errorDescription, anagrup, tip
End If

If identityRs.EOF Then
    identityRs.Close
    Set identityRs = Nothing
    baglanti.RollbackTrans
    On Error GoTo 0
    StopWithError "Yeni ürün numarası alınamadı.", anagrup, tip
End If

yeniID = CLng(identityRs("YeniID"))
identityRs.Close
Set identityRs = Nothing

baglanti.CommitTrans

If Err.Number <> 0 Then
    errorDescription = Err.Description
    Err.Clear
    On Error GoTo 0
    StopWithError "Ürün kaydedildi ancak işlem tamamlanamadı: " & errorDescription, anagrup, tip
End If

On Error GoTo 0

Response.Redirect "edits.asp?id=" & yeniID & "&g=" & anagrup & "&tip=" & tip & "&new=1"
Response.End
%>
