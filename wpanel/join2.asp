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

Dim isim, tip, anagrup, surface
Dim notlar, delivery, yayin, kodu
Dim stok, keyw, descr
Dim sql1

If UCase(Request.ServerVariables("REQUEST_METHOD")) <> "POST" Then
    Response.Redirect "join.asp"
    Response.End
End If

isim = Trim(Request.Form("isim") & "")
tip = Trim(Request.Form("tip") & "")
anagrup = Trim(Request.Form("anagrup") & "")
surface = Trim(Request.Form("surface") & "")
notlar = Request.Form("notlar") & ""
delivery = Trim(Request.Form("delivery") & "")
yayin = Trim(Request.Form("yayin") & "")
kodu = Trim(Request.Form("kodu") & "")
keyw = Request.Form("keyw") & ""
descr = Request.Form("descr") & ""

stok = 10

If isim = "" Then
    Response.Redirect "join.asp?err=isim"
    Response.End
End If

If Not IsNumeric(tip) Then
    Response.Redirect "join.asp?err=tip"
    Response.End
End If

If Not IsNumeric(anagrup) Then
    Response.Redirect "join.asp?err=anagrup"
    Response.End
End If

If Not IsNumeric(surface) Then surface = 0
If Not IsNumeric(delivery) Then delivery = 3
If Not IsNumeric(yayin) Then yayin = 0

sql1 = "INSERT INTO products " & _
       "(isim, tip, anagrup, surface, notlar, delivery, yayin, kodu, stok, keyw, descr) " & _
       "VALUES (" & _
       "'" & SqlText(isim) & "', " & _
       CLng(tip) & ", " & _
       CLng(anagrup) & ", " & _
       CLng(surface) & ", " & _
       "'" & SqlText(notlar) & "', " & _
       CLng(delivery) & ", " & _
       CLng(yayin) & ", " & _
       "'" & SqlText(kodu) & "', " & _
       CLng(stok) & ", " & _
       "'" & SqlText(keyw) & "', " & _
       "'" & SqlText(descr) & "'" & _
       ")"

On Error Resume Next
baglanti.Execute sql1

If Err.Number <> 0 Then
    Response.Write "<h3>Veritabanı hatası</h3>"
    Response.Write "<b>Numara:</b> " & Err.Number & "<br>"
    Response.Write "<b>Açıklama:</b> " & Server.HTMLEncode(Err.Description) & "<br><br>"
    Response.Write "<b>SQL:</b><pre>" & Server.HTMLEncode(sql1) & "</pre>"
    Response.End
End If

On Error GoTo 0

Dim object3, sql
Set object3 = Server.CreateObject("ADODB.Recordset")
sql = "SELECT TOP 1 AffiliateID FROM products ORDER BY AffiliateID DESC"
object3.Open sql, baglanti, 1, 1
%>

<center>
    <br><br>
    <b>Ürün Bilgisi Kaydedildi.</b>
    <br><br>

    <% If Not object3.EOF Then %>
        <a href="edits.asp?id=<%=object3("AffiliateID")%>&g=<%=CLng(anagrup)%>&tip=<%=CLng(tip)%>">
            Görsel Eklemek İçin Tıklayın
        </a>
    <% End If %>
</center>

<%
If object3.State = 1 Then object3.Close
Set object3 = Nothing
%>
