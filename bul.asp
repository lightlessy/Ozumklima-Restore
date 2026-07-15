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
    result = Replace(result, "[", " ")
    result = Replace(result, "]", " ")
    result = Replace(result, "{", " ")
    result = Replace(result, "}", " ")
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
    ContainsSql = "LCase(Nz(" & fieldExpression & ",'')) LIKE '%" & SqlText(LCase(value)) & "%'"
End Function

Function ExactSql(ByVal fieldExpression, ByVal value)
    ExactSql = "LCase(Trim(Nz(" & fieldExpression & ",'')))='" & SqlText(LCase(value)) & "'"
End Function

Function StartsSql(ByVal fieldExpression, ByVal value)
    StartsSql = "LCase(Trim(Nz(" & fieldExpression & ",''))) LIKE '" & SqlText(LCase(value)) & "%'"
End Function

Function KeywordWordSql(ByVal fieldExpression, ByVal value)
    Dim normalizedField

    normalizedField = "LCase(Nz(" & fieldExpression & ",''))"
    normalizedField = "Replace(" & normalizedField & ",',',' ')"
    normalizedField = "Replace(" & normalizedField & ",'.',' ')"
    normalizedField = "Replace(" & normalizedField & ",';',' ')"
    normalizedField = "Replace(" & normalizedField & ",':',' ')"
    normalizedField = "Replace(" & normalizedField & ",'-',' ')"
    normalizedField = "Replace(" & normalizedField & ",'/',' ')"

    KeywordWordSql = "(' ' & " & normalizedField & " & ' ') LIKE '% " & SqlText(LCase(value)) & " %'"
End Function

Dim t, s, d, bul, normalizedQuery, rawParts
Dim tokens(), tokenCount, partIndex, tokenValue, duplicateToken, existingIndex
Dim filterSql, tokenFilterSql, strongMatchSql, scoreSql, tokenScoreSql
Dim sortSql, threshold, sqlQuery, object

t = Trim(Request.QueryString("t") & "")
s = Trim(Request.QueryString("s") & "")
d = Trim(Request.QueryString("d") & "")
bul = Trim(Request.QueryString("bul") & "")

If t <> "" And Not IsNumeric(t) Then t = ""
If s <> "" And Not IsNumeric(s) Then s = ""
If t <> "" Then t = CStr(CLng(t))
If s <> "" Then s = CStr(CLng(s))

normalizedQuery = NormalizeSearch(bul)
tokenCount = 0
ReDim tokens(0)

If normalizedQuery <> "" Then
    rawParts = Split(normalizedQuery, " ")

    For partIndex = 0 To UBound(rawParts)
        tokenValue = Trim(rawParts(partIndex))

        If Len(tokenValue) >= 2 Then
            duplicateToken = False

            For existingIndex = 0 To tokenCount - 1
                If tokens(existingIndex) = tokenValue Then duplicateToken = True
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

Select Case d
    Case "19"
        sortSql = "fiyat ASC, relevance DESC, sira DESC"
    Case "91"
        sortSql = "fiyat DESC, relevance DESC, sira DESC"
    Case Else
        d = "0"
        sortSql = "relevance DESC, sira DESC, AffiliateID DESC"
End Select

filterSql = "p.yayin=1"
If t <> "" Then filterSql = filterSql & " AND p.tip=" & CLng(t)
If s <> "" Then filterSql = filterSql & " AND p.surface=" & CLng(s)

scoreSql = "0"
tokenFilterSql = ""
strongMatchSql = ""

If tokenCount > 0 Then
    scoreSql = scoreSql & _
        "+IIF(" & ExactSql("p.isim", normalizedQuery) & ",1000,0)" & _
        "+IIF(" & StartsSql("p.isim", normalizedQuery) & ",500,0)" & _
        "+IIF(" & ContainsSql("p.isim", normalizedQuery) & ",260,0)" & _
        "+IIF(" & ExactSql("p.kodu", normalizedQuery) & ",900,0)" & _
        "+IIF(" & ExactSql("ag.isim", normalizedQuery) & ",240,0)" & _
        "+IIF(" & ExactSql("tp.isim", normalizedQuery) & ",240,0)" & _
        "+IIF(" & ExactSql("sf.isim", normalizedQuery) & ",220,0)"

    For partIndex = 0 To tokenCount - 1
        tokenValue = tokens(partIndex)

        Dim nameMatch, codeMatch, categoryMatch, keywordMatch, descriptionMatch, anyTokenMatch
        nameMatch = ContainsSql("p.isim", tokenValue)
        codeMatch = ContainsSql("p.kodu", tokenValue)
        categoryMatch = "(" & ContainsSql("ag.isim", tokenValue) & " OR " & ContainsSql("tp.isim", tokenValue) & " OR " & ContainsSql("sf.isim", tokenValue) & ")"
        keywordMatch = KeywordWordSql("p.keyw", tokenValue)
        descriptionMatch = ContainsSql("p.descr", tokenValue)
        anyTokenMatch = "(" & nameMatch & " OR " & codeMatch & " OR " & categoryMatch & " OR " & keywordMatch & " OR " & descriptionMatch & ")"

        If tokenFilterSql <> "" Then tokenFilterSql = tokenFilterSql & " AND "
        tokenFilterSql = tokenFilterSql & anyTokenMatch

        If strongMatchSql <> "" Then strongMatchSql = strongMatchSql & " OR "
        strongMatchSql = strongMatchSql & "(" & nameMatch & " OR " & codeMatch & " OR " & categoryMatch & " OR " & keywordMatch & ")"

        tokenScoreSql = _
            "+IIF(" & nameMatch & ",120,0)" & _
            "+IIF(" & codeMatch & ",150,0)" & _
            "+IIF(" & categoryMatch & ",90,0)" & _
            "+IIF(" & keywordMatch & ",35,0)" & _
            "+IIF(" & descriptionMatch & ",8,0)"

        scoreSql = scoreSql & tokenScoreSql
    Next

    filterSql = filterSql & " AND (" & tokenFilterSql & ") AND (" & strongMatchSql & ")"
End If

threshold = 60
If tokenCount > 1 Then threshold = 60 + ((tokenCount - 1) * 25)

sqlQuery = _
    "SELECT * FROM (" & _
        "SELECT p.*, (" & scoreSql & ") AS relevance " & _
        "FROM ((products AS p " & _
        "LEFT JOIN anagrup AS ag ON p.anagrup=ag.id) " & _
        "LEFT JOIN tip AS tp ON p.tip=tp.id) " & _
        "LEFT JOIN surface AS sf ON p.surface=sf.id " & _
        "WHERE " & filterSql & _
    ") AS ranked " & _
    "WHERE relevance >= " & threshold & " " & _
    "ORDER BY " & sortSql
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
    <style>
        .search-page,.search-page *{box-sizing:border-box}
        .search-page{width:calc(100% - 32px);max-width:1200px;margin:28px auto 65px;font-family:'Source Sans Pro',Arial,sans-serif;color:#333}
        .search-head{display:flex;align-items:center;justify-content:space-between;gap:18px;padding:18px 20px;border:1px solid #e2e2e2;background:#f7f7f7}
        .search-title{margin:0;font-size:1.45rem;line-height:1.25}
        .search-title strong{color:#111}
        .search-sort{min-width:210px;padding:9px 10px;border:1px solid #d4d4d4;background:#fff;font:inherit}
        .search-note{margin:12px 0 0;color:#666;font-size:.94rem}
        .search-empty{padding:46px 22px;margin-top:22px;border:1px dashed #ccc;background:#fafafa;text-align:center;font-size:1.08rem;color:#555}
        .search-grid{display:grid;grid-template-columns:repeat(3,minmax(0,1fr));gap:22px;margin-top:25px}
        .search-card{min-width:0;border:1px solid #e3e3e3;background:#fff;transition:border-color .2s ease,transform .2s ease,box-shadow .2s ease}
        .search-card:hover{border-color:#1f9bb7;transform:translateY(-2px);box-shadow:0 7px 20px rgba(0,0,0,.07)}
        .search-card a{display:flex;height:100%;flex-direction:column;color:#222;text-decoration:none}
        .search-card__image{display:flex;aspect-ratio:4/3;align-items:center;justify-content:center;overflow:hidden;background:#f5f5f5}
        .search-card__image img{display:block;width:100%;height:100%;object-fit:contain}
        .search-card__placeholder{padding:24px;text-align:center;color:#777}
        .search-card__body{display:flex;flex:1;flex-direction:column;padding:17px}
        .search-card__name{font-size:1.18rem;font-weight:700;line-height:1.35;overflow-wrap:anywhere}
        .search-card__description{margin-top:9px;color:#666;line-height:1.45;display:-webkit-box;-webkit-line-clamp:3;-webkit-box-orient:vertical;overflow:hidden}
        .search-card__action{align-self:flex-end;margin-top:auto;padding-top:17px;font-weight:700;color:#147d94}
        @media(max-width:900px){.search-grid{grid-template-columns:repeat(2,minmax(0,1fr))}}
        @media(max-width:600px){.search-page{width:calc(100% - 20px);margin-top:18px}.search-head{align-items:stretch;flex-direction:column;padding:15px}.search-sort{width:100%;min-width:0}.search-grid{grid-template-columns:1fr;gap:15px}.search-card a{display:grid;grid-template-columns:120px minmax(0,1fr)}.search-card__image{aspect-ratio:1}.search-card__body{padding:14px}.search-card__description{font-size:.93rem;-webkit-line-clamp:2}}
    </style>
</head>
<body>
<!--#include file='ust.asp'-->

<main class="search-page">
    <div class="search-head">
        <h1 class="search-title">
        <% If bul <> "" Then %>
            Arama: <strong><%=Server.HTMLEncode(bul)%></strong>
        <% Else %>
            Ürün arama
        <% End If %>
        </h1>

    <% If tokenCount > 0 Then %>
        <select class="search-sort" id="searchSort" aria-label="Arama sonuçlarını sırala">
            <option value="?bul=<%=Server.URLEncode(bul)%>&amp;t=<%=Server.URLEncode(t)%>&amp;s=<%=Server.URLEncode(s)%>&amp;d=0"<% If d = "0" Then Response.Write " selected" %>>En alakalı</option>
            <option value="?bul=<%=Server.URLEncode(bul)%>&amp;t=<%=Server.URLEncode(t)%>&amp;s=<%=Server.URLEncode(s)%>&amp;d=19"<% If d = "19" Then Response.Write " selected" %>>Ucuzdan pahalıya</option>
            <option value="?bul=<%=Server.URLEncode(bul)%>&amp;t=<%=Server.URLEncode(t)%>&amp;s=<%=Server.URLEncode(s)%>&amp;d=91"<% If d = "91" Then Response.Write " selected" %>>Pahalıdan ucuza</option>
        </select>
    <% End If %>
    </div>

<% If tokenCount = 0 Then %>
    <div class="search-empty">Arama yapmak için en az iki karakterlik bir ürün adı, model, kategori veya özellik yazın.</div>
<% Else %>
    <p class="search-note">Sonuçlar ürün adı, model kodu, kategori, ürün grubu ve kontrollü anahtar kelime eşleşmesine göre sıralanır.</p>
<%
    Set object = Server.CreateObject("ADODB.Recordset")
    object.Open sqlQuery, baglanti, 1, 1

    If object.EOF Then
%>
    <div class="search-empty">“<%=Server.HTMLEncode(bul)%>” için yeterince alakalı bir ürün bulunamadı.</div>
<%
    Else
%>
    <div class="search-grid">
<%
        Do While Not object.EOF
            Dim productName, productId, productImage, productLink, productDescription
            productName = object("isim") & ""
            productId = object("AffiliateID") & ""
            productImage = object("foto1") & ""
            productDescription = object("descr") & ""
            productLink = "detay.asp?id=" & productId

            If productName <> "" Then productLink = cevir(productName) & "-" & productId
%>
        <article class="search-card">
            <a href="<%=Server.HTMLEncode(productLink)%>">
                <div class="search-card__image">
                <% If productImage <> "" Then %>
                    <img src="urunler/<%=Server.HTMLEncode(productImage)%>" width="480" height="360" loading="lazy" alt="<%=Server.HTMLEncode(productName)%>">
                <% Else %>
                    <span class="search-card__placeholder">Görsel yakında eklenecek</span>
                <% End If %>
                </div>
                <div class="search-card__body">
                    <div class="search-card__name"><%=Server.HTMLEncode(productName)%></div>
                <% If Trim(productDescription) <> "" Then %>
                    <div class="search-card__description"><%=Server.HTMLEncode(productDescription)%></div>
                <% End If %>
                    <div class="search-card__action">İncele →</div>
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

    If object.State = 1 Then object.Close
    Set object = Nothing
End If
%>
</main>

<!--#include file='alt.asp'-->
<script>
$(function(){
    $('#searchSort').on('change',function(){
        var target=$(this).val();
        if(target){window.location.href=target;}
    });
});
</script>
</body>
</html>