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
    KeywordSql = ContainsSql("p.keyw", value)
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

whereSql = "(p.yayin<>0)"
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
            KeywordSql(tokenValue) & " OR " & _
            ContainsSql("p.descr", tokenValue) & _
        ")"

        If tokenWhere <> "" Then tokenWhere = tokenWhere & " OR "
        tokenWhere = tokenWhere & tokenCandidate

        tokenScore = _
            "+IIF(" & ContainsSql("p.isim", tokenValue) & ",120,0)" & _
            "+IIF(" & ContainsSql("p.kodu", tokenValue) & ",150,0)" & _
            "+IIF((" & ContainsSql("ag.isim", tokenValue) & _
                " OR " & ContainsSql("tp.isim", tokenValue) & _
                " OR " & ContainsSql("sf.isim", tokenValue) & "),90,0)" & _
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
    <style>
        .search-page,
        .search-page * {
            box-sizing: border-box;
        }

        .search-page {
            width: min(1240px, calc(100% - 40px));
            margin: 34px auto 72px;
            font-family: 'Source Sans Pro', Arial, sans-serif;
            color: #1f2933;
        }

        .search-head {
            display: grid;
            grid-template-columns: minmax(0, 1fr) auto;
            align-items: center;
            gap: 24px;
            padding: 24px 26px;
            border: 1px solid #e4e7eb;
            border-radius: 14px;
            background: linear-gradient(180deg, #ffffff 0%, #f8fafb 100%);
            box-shadow: 0 10px 30px rgba(15, 23, 42, .05);
        }

        .search-title {
            min-width: 0;
            margin: 0;
            font-size: clamp(1.35rem, 2vw, 1.85rem);
            line-height: 1.2;
            letter-spacing: -.02em;
            overflow-wrap: anywhere;
        }

        .search-title strong {
            color: #147d94;
        }

        .search-sort {
            width: 230px;
            max-width: 100%;
            padding: 12px 42px 12px 14px;
            border: 1px solid #d8dee4;
            border-radius: 10px;
            background: #fff;
            color: #202b33;
            font: inherit;
            cursor: pointer;
            outline: none;
        }

        .search-sort:focus {
            border-color: #147d94;
            box-shadow: 0 0 0 3px rgba(20, 125, 148, .14);
        }

        .search-note {
            margin: 18px 2px 0;
            color: #66727c;
            font-size: .96rem;
            line-height: 1.55;
        }

        .search-empty {
            margin-top: 24px;
            padding: 54px 24px;
            border: 1px dashed #cdd5dc;
            border-radius: 14px;
            background: #fafcfd;
            text-align: center;
            color: #52606d;
            font-size: 1.06rem;
        }

        .search-grid {
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(240px, 1fr));
            gap: 22px;
            margin-top: 26px;
            align-items: stretch;
        }

        .search-card {
            min-width: 0;
            overflow: hidden;
            border: 1px solid #e2e7eb;
            border-radius: 14px;
            background: #fff;
            box-shadow: 0 8px 24px rgba(15, 23, 42, .055);
            transition: transform .22s ease, box-shadow .22s ease, border-color .22s ease;
        }

        .search-card:hover {
            transform: translateY(-4px);
            border-color: #c9d7dc;
            box-shadow: 0 16px 34px rgba(15, 23, 42, .11);
        }

        .search-card a {
            display: flex;
            height: 100%;
            min-height: 100%;
            flex-direction: column;
            color: inherit;
            text-decoration: none;
        }

        .search-card__image {
            display: flex;
            width: 100%;
            aspect-ratio: 16 / 11;
            align-items: center;
            justify-content: center;
            overflow: hidden;
            border-bottom: 1px solid #edf0f2;
            background: linear-gradient(180deg, #fbfbfb 0%, #f3f5f6 100%);
        }

        .search-card__image img {
            display: block;
            width: 100%;
            height: 100%;
            padding: 14px;
            object-fit: contain;
            transition: transform .28s ease;
        }

        .search-card:hover .search-card__image img {
            transform: scale(1.035);
        }

        .search-card__placeholder {
            padding: 24px;
            text-align: center;
            color: #7b8794;
        }

        .search-card__body {
            display: flex;
            min-height: 210px;
            flex: 1;
            flex-direction: column;
            padding: 19px 19px 18px;
        }

        .search-card__name {
            display: -webkit-box;
            overflow: hidden;
            color: #1f2933;
            font-size: 1.08rem;
            font-weight: 700;
            line-height: 1.42;
            overflow-wrap: anywhere;
            -webkit-box-orient: vertical;
            -webkit-line-clamp: 3;
        }

        .search-card__description {
            display: -webkit-box;
            overflow: hidden;
            margin-top: 11px;
            color: #66727c;
            font-size: .96rem;
            line-height: 1.52;
            -webkit-box-orient: vertical;
            -webkit-line-clamp: 3;
        }

        .search-card__action {
            display: inline-flex;
            align-items: center;
            gap: 8px;
            align-self: flex-start;
            margin-top: auto;
            padding-top: 19px;
            color: #147d94;
            font-weight: 700;
        }

        .search-card__action::after {
            content: "→";
            transition: transform .2s ease;
        }

        .search-card:hover .search-card__action::after {
            transform: translateX(4px);
        }

        @media (min-width: 1180px) {
            .search-grid {
                grid-template-columns: repeat(4, minmax(0, 1fr));
            }
        }

        @media (max-width: 820px) {
            .search-page {
                width: min(100% - 28px, 760px);
                margin-top: 24px;
            }

            .search-head {
                grid-template-columns: 1fr;
                gap: 16px;
                padding: 20px;
            }

            .search-sort {
                width: 100%;
            }

            .search-grid {
                grid-template-columns: repeat(2, minmax(0, 1fr));
                gap: 16px;
            }
        }

        @media (max-width: 560px) {
            .search-page {
                width: calc(100% - 20px);
                margin: 16px auto 54px;
            }

            .search-head {
                border-radius: 11px;
                padding: 17px;
            }

            .search-grid {
                grid-template-columns: 1fr;
                gap: 14px;
                margin-top: 20px;
            }

            .search-card a {
                display: grid;
                grid-template-columns: 132px minmax(0, 1fr);
                min-height: 150px;
            }

            .search-card__image {
                height: 100%;
                min-height: 150px;
                aspect-ratio: auto;
                border-right: 1px solid #edf0f2;
                border-bottom: 0;
            }

            .search-card__image img {
                padding: 10px;
            }

            .search-card__body {
                min-height: 0;
                padding: 15px 14px;
            }

            .search-card__name {
                font-size: 1rem;
                -webkit-line-clamp: 3;
            }

            .search-card__description {
                margin-top: 7px;
                font-size: .9rem;
                -webkit-line-clamp: 2;
            }

            .search-card__action {
                padding-top: 12px;
                font-size: .94rem;
            }
        }

        @media (max-width: 390px) {
            .search-card a {
                grid-template-columns: 112px minmax(0, 1fr);
            }
        }
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
                <option value="?bul=<%=Server.URLEncode(bul)%>&amp;t=<%=Server.URLEncode(t)%>&amp;s=<%=Server.URLEncode(s)%>&amp;d=0"<% If d="0" Then Response.Write " selected" %>>En alakalı</option>
                <option value="?bul=<%=Server.URLEncode(bul)%>&amp;t=<%=Server.URLEncode(t)%>&amp;s=<%=Server.URLEncode(s)%>&amp;d=19"<% If d="19" Then Response.Write " selected" %>>Ucuzdan pahalıya</option>
                <option value="?bul=<%=Server.URLEncode(bul)%>&amp;t=<%=Server.URLEncode(t)%>&amp;s=<%=Server.URLEncode(s)%>&amp;d=91"<% If d="91" Then Response.Write " selected" %>>Pahalıdan ucuza</option>
            </select>
        <% End If %>
    </div>

    <% If tokenCount = 0 Then %>
        <div class="search-empty">Arama yapmak için en az iki karakterlik bir ürün adı, model, kategori veya özellik yazın.</div>
    <% Else %>
<%
Set object = Server.CreateObject("ADODB.Recordset")

On Error Resume Next
object.Open sqlQuery, baglanti, 0, 1

If Err.Number <> 0 Then
    Dim sqlErrorNumber, sqlErrorDescription

    sqlErrorNumber = Err.Number
    sqlErrorDescription = Err.Description

    Err.Clear
    On Error GoTo 0

    Response.Write "<div class='search-empty'>Arama şu anda tamamlanamadı. Lütfen tekrar deneyin.</div>"

    Set object = Nothing
    Response.End
End If

On Error GoTo 0

If object.EOF Then
%>
        <div class="search-empty">“<%=Server.HTMLEncode(bul)%>” için alakalı ürün bulunamadı.</div>
<% Else %>
        <p class="search-note">Sonuçlar ürün adı, model kodu, kategori, yüzey tipi, açıklama ve anahtar kelime eşleşmesine göre sıralanır.</p>

        <div class="search-grid">
<%
Do While Not object.EOF
    Dim urunIsim, urunID, urunFoto, urunLink, urunDescr

    urunIsim = object("isim") & ""
    urunID = object("AffiliateID") & ""
    urunFoto = object("foto1") & ""
    urunDescr = object("descr") & ""
    urunLink = cevir(urunIsim) & "-" & urunID
%>
            <article class="search-card">
                <a href="<%=Server.HTMLEncode(urunLink)%>">
                    <div class="search-card__image">
                        <% If urunFoto <> "" Then %>
                            <img src="urunler/<%=Server.HTMLEncode(urunFoto)%>" alt="<%=Server.HTMLEncode(urunIsim)%>" loading="lazy">
                        <% Else %>
                            <span class="search-card__placeholder">Görsel yakında eklenecek</span>
                        <% End If %>
                    </div>

                    <div class="search-card__body">
                        <div class="search-card__name"><%=Server.HTMLEncode(urunIsim)%></div>

                        <% If Trim(urunDescr) <> "" Then %>
                            <div class="search-card__description"><%=Server.HTMLEncode(urunDescr)%></div>
                        <% End If %>

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
%>
    <% End If %>
</main>
<!--#include file='alt.asp'-->
<script>
$(function () {
    $('#searchSort').on('change', function () {
        if (this.value) {
            window.location.href = this.value;
        }
    });
});
</script>
</body>
</html>
