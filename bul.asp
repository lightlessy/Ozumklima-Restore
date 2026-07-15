<%@ Language="VBScript" CodePage="65001" %>
<%
Response.Buffer = True
Response.CodePage = 65001
Session.CodePage = 65001
Session.LCID = 1055
Response.Charset = "utf-8"
%>
<!--#include file='db.asp'-->
<%
Function SqlText(ByVal value)
    SqlText = Replace(CStr(value & ""), "'", "''")
End Function

Function NormalizeSearch(ByVal value)
    Dim result
    result = LCase(Trim(CStr(value & "")))
    result = Replace(result, vbTab, " ")
    result = Replace(result, vbCr, " ")
    result = Replace(result, vbLf, " ")
    result = Replace(result, ",", " ")
    result = Replace(result, ".", " ")
    result = Replace(result, ";", " ")
    result = Replace(result, ":", " ")
    result = Replace(result, "!", " ")
    result = Replace(result, "?", " ")
    result = Replace(result, "(", " ")
    result = Replace(result, ")", " ")
    result = Replace(result, "-", " ")
    result = Replace(result, "_", " ")
    result = Replace(result, "/", " ")
    result = Replace(result, "\", " ")
    result = Replace(result, Chr(34), " ")
    result = Replace(result, "'", " ")

    Do While InStr(result, "  ") > 0
        result = Replace(result, "  ", " ")
    Loop

    NormalizeSearch = Trim(result)
End Function

Function ContainsSql(ByVal fieldExpression, ByVal value)
    ContainsSql = "LCase(IIF(" & fieldExpression & " Is Null,''," & fieldExpression & ")) LIKE '%" & SqlText(LCase(value)) & "%'"
End Function

Function ExactSql(ByVal fieldExpression, ByVal value)
    ExactSql = "LCase(Trim(IIF(" & fieldExpression & " Is Null,''," & fieldExpression & ")))='" & SqlText(LCase(value)) & "'"
End Function

Function StartsSql(ByVal fieldExpression, ByVal value)
    StartsSql = "LCase(Trim(IIF(" & fieldExpression & " Is Null,''," & fieldExpression & "))) LIKE '" & SqlText(LCase(value)) & "%'"
End Function

Function KeywordSql(ByVal value)
    Dim fieldValue
    fieldValue = "LCase(IIF(p.keyw Is Null,'',p.keyw))"
    fieldValue = "Replace(" & fieldValue & ",',',' ')"
    fieldValue = "Replace(" & fieldValue & ",';',' ')"
    fieldValue = "Replace(" & fieldValue & ",'.',' ')"
    fieldValue = "Replace(" & fieldValue & ",'-',' ')"
    KeywordSql = "(' ' & " & fieldValue & " & ' ') LIKE '% " & SqlText(LCase(value)) & " %'"
End Function

Dim t, s, d, bul, normalizedQuery, rawParts
Dim tokens(), tokenCount, i, j, tokenValue, duplicateToken
Dim whereSql, tokenWhere, tokenCandidate, relevanceSql, tokenScore
Dim sortSql, sqlQuery, object

t = Trim(Request("t") & "")
s = Trim(Request("s") & "")
d = Trim(Request("d") & "")
bul = Trim(Request("bul") & "")

If t <> "" And Not IsNumeric(t) Then t = ""
If s <> "" And Not IsNumeric(s) Then s = ""
If t <> "" Then t = CStr(CLng(t))
If s <> "" Then s = CStr(CLng(s))

normalizedQuery = NormalizeSearch(bul)
tokenCount = 0
ReDim tokens(0)

If normalizedQuery <> "" Then
    rawParts = Split(normalizedQuery, " ")

    For i = 0 To UBound(rawParts)
        tokenValue = Trim(rawParts(i))

        If Len(tokenValue) >= 2 Then
            duplicateToken = False

            For j = 0 To tokenCount - 1
                If tokens(j) = tokenValue Then duplicateToken = True
            Next

            If Not duplicateToken Then
                If tokenCount > 0 Then ReDim Preserve tokens(tokenCount)
                tokens(tokenCount) = tokenValue
                tokenCount = tokenCount + 1
                If tokenCount = 6 Then Exit For
            End If
        End If
    Next
End If

whereSql = "p.yayin=1"
If t <> "" Then whereSql = whereSql & " AND p.tip=" & CLng(t)
If s <> "" Then whereSql = whereSql & " AND p.surface=" & CLng(s)

relevanceSql = "0"
tokenWhere = ""

If tokenCount > 0 Then
    relevanceSql = relevanceSql & _
        "+IIF(" & ExactSql("p.isim", normalizedQuery) & ",1000,0)" & _
        "+IIF(" & ExactSql("p.kodu", normalizedQuery) & ",900,0)" & _
        "+IIF(" & StartsSql("p.isim", normalizedQuery) & ",500,0)" & _
        "+IIF(" & ContainsSql("p.isim", normalizedQuery) & ",260,0)"

    For i = 0 To tokenCount - 1
        tokenValue = tokens(i)

        tokenCandidate = "(" & _
            ContainsSql("p.isim", tokenValue) & " OR " & _
            ContainsSql("p.kodu", tokenValue) & " OR " & _
            ContainsSql("ag.isim", tokenValue) & " OR " & _
            ContainsSql("tp.isim", tokenValue) & " OR " & _
            ContainsSql("sf.isim", tokenValue) & " OR " & _
            KeywordSql(tokenValue) & _
        ")"

        If tokenWhere <> "" Then tokenWhere = tokenWhere & " AND "
        tokenWhere = tokenWhere & tokenCandidate

        tokenScore = _
            "+IIF(" & ContainsSql("p.isim", tokenValue) & ",120,0)" & _
            "+IIF(" & ContainsSql("p.kodu", tokenValue) & ",150,0)" & _
            "+IIF((" & ContainsSql("ag.isim", tokenValue) & " OR " & ContainsSql("tp.isim", tokenValue) & " OR " & ContainsSql("sf.isim", tokenValue) & "),90,0)" & _
            "+IIF(" & KeywordSql(tokenValue) & ",35,0)" & _
            "+IIF(" & ContainsSql("p.descr", tokenValue) & ",8,0)"

        relevanceSql = relevanceSql & tokenScore
    Next

    whereSql = whereSql & " AND (" & tokenWhere & ")"
End If

Select Case d
    Case "19"
        sortSql = "p.fiyat ASC, (" & relevanceSql & ") DESC, p.sira DESC"
    Case "91"
        sortSql = "p.fiyat DESC, (" & relevanceSql & ") DESC, p.sira DESC"
    Case Else
        d = "0"
        sortSql = "(" & relevanceSql & ") DESC, p.sira DESC, p.AffiliateID DESC"
End Select

sqlQuery = _
    "SELECT p.* FROM ((products AS p " & _
    "LEFT JOIN anagrup AS ag ON p.anagrup=ag.id) " & _
    "LEFT JOIN tip AS tp ON p.tip=tp.id) " & _
    "LEFT JOIN surface AS sf ON p.surface=sf.id " & _
    "WHERE " & whereSql & " ORDER BY " & sortSql
%>
<!doctype html>
<html lang="tr">
<head>
    <meta charset="utf-8">
    <meta name="viewport" content="width=device-width, initial-scale=1, viewport-fit=cover">
    <meta name="robots" content="noindex,follow">
    <title><% If bul <> "" Then %><%=Server.HTMLEncode(bul)%> Arama Sonuçları<% Else %>Ürün Arama<% End If %> | Özüm Klima</title>
    <link rel="stylesheet" href="css/font.css">
    <link rel="shortcut icon" href="images/favicon.png">
    <link rel="stylesheet" href="css/fontawesome-all.css">
    <script src="js/jquery-2.1.1.js"></script>
</head>
<body>
<!--#include file='ust.asp'-->
<main class="search-page">
    <div class="search-head">
        <h1 class="search-title"><% If bul <> "" Then %>Arama: <strong><%=Server.HTMLEncode(bul)%></strong><% Else %>Ürün arama<% End If %></h1>
    </div>

<% If tokenCount = 0 Then %>
    <div class="search-empty">Arama yapmak için en az iki karakterlik bir ürün adı, model, kategori veya özellik yazın.</div>
<% Else %>
<%
Set object = Server.CreateObject("ADODB.Recordset")
On Error Resume Next
object.Open sqlQuery, baglanti, 1, 1
If Err.Number <> 0 Then
    Response.Write "<div class='search-empty'>Arama şu anda tamamlanamadı. Lütfen tekrar deneyin.</div>"
    Err.Clear
ElseIf object.EOF Then
%>
    <div class="search-empty">“<%=Server.HTMLEncode(bul)%>” için alakalı ürün bulunamadı.</div>
<% Else %>
    <div class="search-grid">
<%
Do While Not object.EOF
    Dim urunIsim, urunID, urunFoto, urunLink
    urunIsim = object("isim") & ""
    urunID = object("AffiliateID") & ""
    urunFoto = object("foto1") & ""
    urunLink = cevir(urunIsim) & "-" & urunID
%>
        <article class="search-card">
            <a href="<%=Server.HTMLEncode(urunLink)%>">
                <div class="search-card__image">
                <% If urunFoto <> "" Then %><img src="urunler/<%=Server.HTMLEncode(urunFoto)%>" alt="<%=Server.HTMLEncode(urunIsim)%>" loading="lazy"><% Else %><span class="search-card__placeholder">Görsel yakında eklenecek</span><% End If %>
                </div>
                <div class="search-card__body">
                    <div class="search-card__name"><%=Server.HTMLEncode(urunIsim)%></div>
                    <div class="search-card__action">Ürünü incele</div>
                </div>
            </a>
        </article>
<%
    object.MoveNext
Loop
%>
    </div>
<%
End If
If Not object Is Nothing Then
    If object.State = 1 Then object.Close
End If
Set object = Nothing
On Error GoTo 0
%>
<% End If %>
</main>
<!--#include file='alt.asp'-->
</body>
</html>