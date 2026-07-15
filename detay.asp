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
Function PlainText(ByVal value)
    Dim textValue, tagStart, tagEnd

    textValue = CStr(value & "")
    textValue = Replace(textValue, vbCrLf, " ")
    textValue = Replace(textValue, vbCr, " ")
    textValue = Replace(textValue, vbLf, " ")
    textValue = Replace(textValue, "&nbsp;", " ", 1, -1, 1)
    textValue = Replace(textValue, "&amp;", "&", 1, -1, 1)
    textValue = Replace(textValue, "&quot;", Chr(34), 1, -1, 1)
    textValue = Replace(textValue, "&#39;", "'", 1, -1, 1)

    Do While InStr(textValue, "<") > 0
        tagStart = InStr(textValue, "<")
        tagEnd = InStr(tagStart, textValue, ">")

        If tagEnd <= tagStart Then Exit Do
        textValue = Left(textValue, tagStart - 1) & " " & Mid(textValue, tagEnd + 1)
    Loop

    textValue = Trim(textValue)

    Do While InStr(textValue, "  ") > 0
        textValue = Replace(textValue, "  ", " ")
    Loop

    PlainText = textValue
End Function

Function JsonText(ByVal value)
    Dim textValue

    textValue = CStr(value & "")
    textValue = Replace(textValue, "\", "\\")
    textValue = Replace(textValue, Chr(34), "\" & Chr(34))
    textValue = Replace(textValue, vbCrLf, "\n")
    textValue = Replace(textValue, vbCr, "\n")
    textValue = Replace(textValue, vbLf, "\n")
    textValue = Replace(textValue, "<", "\u003c")
    textValue = Replace(textValue, ">", "\u003e")
    textValue = Replace(textValue, "&", "\u0026")

    JsonText = textValue
End Function

Function NumberOrZero(ByVal value)
    If IsNumeric(value) Then
        NumberOrZero = CDbl(value)
    Else
        NumberOrZero = 0
    End If
End Function

Function JsonNumber(ByVal value)
    JsonNumber = Replace(CStr(CDbl(value)), ",", ".")
End Function

Function ProductSlug(ByVal value)
    Dim slugValue, slugRegex

    slugValue = LCase(Trim(CStr(value & "")))
    slugValue = Replace(slugValue, "ç", "c")
    slugValue = Replace(slugValue, "ğ", "g")
    slugValue = Replace(slugValue, "ı", "i")
    slugValue = Replace(slugValue, "ö", "o")
    slugValue = Replace(slugValue, "ş", "s")
    slugValue = Replace(slugValue, "ü", "u")
    slugValue = Replace(slugValue, "Ç", "c")
    slugValue = Replace(slugValue, "Ğ", "g")
    slugValue = Replace(slugValue, "İ", "i")
    slugValue = Replace(slugValue, "I", "i")
    slugValue = Replace(slugValue, "Ö", "o")
    slugValue = Replace(slugValue, "Ş", "s")
    slugValue = Replace(slugValue, "Ü", "u")
    slugValue = Replace(slugValue, "ý", "i")
    slugValue = Replace(slugValue, "ð", "g")
    slugValue = Replace(slugValue, "þ", "s")

    Set slugRegex = New RegExp
    slugRegex.Global = True
    slugRegex.IgnoreCase = True
    slugRegex.Pattern = "[^a-z0-9]+"
    slugValue = slugRegex.Replace(slugValue, "-")
    slugValue = Trim(slugValue)

    Do While Left(slugValue, 1) = "-"
        slugValue = Mid(slugValue, 2)
    Loop

    Do While Right(slugValue, 1) = "-"
        slugValue = Left(slugValue, Len(slugValue) - 1)
    Loop

    If slugValue = "" Then slugValue = "urun"
    ProductSlug = slugValue
End Function

Function IsHexColor(ByVal value)
    Dim colorValue, colorRegex

    colorValue = Replace(Trim(CStr(value & "")), "#", "")
    Set colorRegex = New RegExp
    colorRegex.Global = False
    colorRegex.IgnoreCase = True
    colorRegex.Pattern = "^([0-9a-f]{3}|[0-9a-f]{6})$"
    IsHexColor = colorRegex.Test(colorValue)
End Function

Dim rawQuery, productToken, queryParts, pidi
Dim uruns, anas, tipi, sorf, Sorgula
Dim productFound

rawQuery = Trim(Request.ServerVariables("QUERY_STRING") & "")
productToken = Trim(Request.QueryString("id") & "")

If productToken = "" Then
    productToken = rawQuery

    If InStr(productToken, "&") > 0 Then
        productToken = Left(productToken, InStr(productToken, "&") - 1)
    End If

    If InStrRev(productToken, "/") > 0 Then
        productToken = Mid(productToken, InStrRev(productToken, "/") + 1)
    End If

    queryParts = Split(productToken, "-")
    productToken = queryParts(UBound(queryParts))
End If

pidi = ""
If IsNumeric(productToken) Then
    If CLng(productToken) > 0 Then pidi = CStr(CLng(productToken))
End If

productFound = False
Set uruns = Nothing

If pidi <> "" Then
    Set uruns = Server.CreateObject("ADODB.Recordset")
    Sorgula = "SELECT * FROM products WHERE AffiliateID=" & CLng(pidi) & " AND yayin=1"
    uruns.Open Sorgula, baglanti, 1, 1

    If Not uruns.EOF Then productFound = True
End If

If Not productFound Then
    Response.Status = "404 Not Found"
%>
<!doctype html>
<html lang="tr">
<head>
    <meta charset="utf-8">
    <meta name="viewport" content="width=device-width, initial-scale=1">
    <meta name="robots" content="noindex,follow">
    <title>Ürün bulunamadı | Özüm Klima</title>
    <link rel="stylesheet" href="css/font.css">
    <link rel="shortcut icon" href="images/favicon.png">
    <style>
        .product-not-found{width:calc(100% - 32px);max-width:900px;margin:70px auto;padding:45px 24px;box-sizing:border-box;text-align:center;background:#fff;border:1px solid #e2e2e2;font-family:'Source Sans Pro',Arial,sans-serif;color:#333}
        .product-not-found h1{margin:0 0 12px;font-size:2rem}
        .product-not-found a{display:inline-block;margin-top:16px;padding:11px 18px;background:#3c3c3c;color:#fff;border-radius:5px}
    </style>
</head>
<body>
<!--#include file='ust.asp'-->
<main class="product-not-found">
    <h1>Ürün bulunamadı</h1>
    <p>Aradığınız ürün kaldırılmış, yayından alınmış veya bağlantı hatalı olabilir.</p>
    <a href="/">Ana sayfaya dön</a>
</main>
<!--#include file='alt.asp'-->
</body>
</html>
<%
    If Not (uruns Is Nothing) Then
        If uruns.State = 1 Then uruns.Close
        Set uruns = Nothing
    End If

    Response.End
End If

Dim categoryName, typeName, surfaceName
Dim productName, productCode, productDescription, displayDescription
Dim canonicalUrl, seoTitle, seoDescription, absoluteImageUrl
Dim foto(5), schemaImages, imageIndex, imageCount
Dim productSchema, breadcrumbSchema, categoryPath
Dim regularPrice, discountPrice, effectivePrice, stockCount, deliveryDays
Dim hasColors, hasSizes, optionIndex, optionValue

categoryName = ""
typeName = ""
surfaceName = ""

Set anas = Server.CreateObject("ADODB.Recordset")
anas.Open "SELECT id, isim FROM anagrup WHERE id=" & CLng(uruns("anagrup")), baglanti, 1, 1
If Not anas.EOF Then categoryName = anas("isim") & ""

Set tipi = Server.CreateObject("ADODB.Recordset")
tipi.Open "SELECT id, isim FROM tip WHERE id=" & CLng(uruns("tip")), baglanti, 1, 1
If Not tipi.EOF Then typeName = tipi("isim") & ""

Set sorf = Nothing
If IsNumeric(uruns("surface") & "") Then
    If CLng(uruns("surface")) > 0 Then
        Set sorf = Server.CreateObject("ADODB.Recordset")
        sorf.Open "SELECT id, isim FROM surface WHERE id=" & CLng(uruns("surface")), baglanti, 1, 1
        If Not sorf.EOF Then surfaceName = sorf("isim") & ""
    End If
End If

productName = Trim(uruns("isim") & "")
productCode = Trim(uruns("kodu") & "")
productDescription = PlainText(uruns("descr") & "")
displayDescription = productDescription

If displayDescription = "" Then
    displayDescription = productName & " için Antalya'da ücretsiz keşif, profesyonel montaj ve teknik servis desteği alın."
End If

seoDescription = displayDescription
If Len(seoDescription) > 157 Then seoDescription = Left(seoDescription, 154) & "..."

seoTitle = productName & " | Özüm Klima"
If Len(seoTitle) > 65 Then seoTitle = Left(productName, 47) & "... | Özüm Klima"

canonicalUrl = "https://ozumklima.com/" & ProductSlug(productName) & "-" & CLng(uruns("AffiliateID"))

schemaImages = ""
imageCount = 0
For imageIndex = 1 To 5
    foto(imageIndex) = Trim(uruns("foto" & imageIndex) & "")

    If foto(imageIndex) <> "" Then
        absoluteImageUrl = "https://ozumklima.com/urunler/" & foto(imageIndex)
        imageCount = imageCount + 1

        If schemaImages <> "" Then schemaImages = schemaImages & ","
        schemaImages = schemaImages & Chr(34) & JsonText(absoluteImageUrl) & Chr(34)
    End If
Next

regularPrice = NumberOrZero(uruns("fiyat"))
discountPrice = NumberOrZero(uruns("indirim"))
stockCount = NumberOrZero(uruns("stok"))
deliveryDays = NumberOrZero(uruns("delivery"))
effectivePrice = regularPrice
If discountPrice > 0 Then effectivePrice = discountPrice

hasColors = False
hasSizes = False
For optionIndex = 1 To 10
    optionValue = Trim(uruns("renk" & optionIndex) & "")
    If optionValue <> "" And optionValue <> "0" Then hasColors = True

    optionValue = Trim(uruns("beden" & optionIndex) & "")
    If optionValue <> "" And optionValue <> "0" Then hasSizes = True
Next

categoryPath = categoryName
If typeName <> "" Then
    If categoryPath <> "" Then categoryPath = categoryPath & " > "
    categoryPath = categoryPath & typeName
End If
If surfaceName <> "" Then
    If categoryPath <> "" Then categoryPath = categoryPath & " > "
    categoryPath = categoryPath & surfaceName
End If

productSchema = "{" & _
    Chr(34) & "@context" & Chr(34) & ":" & Chr(34) & "https://schema.org" & Chr(34) & "," & _
    Chr(34) & "@type" & Chr(34) & ":" & Chr(34) & "Product" & Chr(34) & "," & _
    Chr(34) & "@id" & Chr(34) & ":" & Chr(34) & JsonText(canonicalUrl & "#product") & Chr(34) & "," & _
    Chr(34) & "url" & Chr(34) & ":" & Chr(34) & JsonText(canonicalUrl) & Chr(34) & "," & _
    Chr(34) & "name" & Chr(34) & ":" & Chr(34) & JsonText(productName) & Chr(34) & "," & _
    Chr(34) & "description" & Chr(34) & ":" & Chr(34) & JsonText(displayDescription) & Chr(34)

If schemaImages <> "" Then
    productSchema = productSchema & "," & Chr(34) & "image" & Chr(34) & ":[" & schemaImages & "]"
End If

If productCode <> "" And productCode <> "0" Then
    productSchema = productSchema & "," & Chr(34) & "sku" & Chr(34) & ":" & Chr(34) & JsonText(productCode) & Chr(34)
End If

If categoryPath <> "" Then
    productSchema = productSchema & "," & Chr(34) & "category" & Chr(34) & ":" & Chr(34) & JsonText(categoryPath) & Chr(34)
End If

If effectivePrice > 0 Then
    productSchema = productSchema & "," & Chr(34) & "offers" & Chr(34) & ":{" & _
        Chr(34) & "@type" & Chr(34) & ":" & Chr(34) & "Offer" & Chr(34) & "," & _
        Chr(34) & "url" & Chr(34) & ":" & Chr(34) & JsonText(canonicalUrl) & Chr(34) & "," & _
        Chr(34) & "priceCurrency" & Chr(34) & ":" & Chr(34) & "TRY" & Chr(34) & "," & _
        Chr(34) & "price" & Chr(34) & ":" & Chr(34) & JsonNumber(effectivePrice) & Chr(34) & "," & _
        Chr(34) & "availability" & Chr(34) & ":" & Chr(34)

    If stockCount > 0 Then
        productSchema = productSchema & "https://schema.org/InStock"
    Else
        productSchema = productSchema & "https://schema.org/OutOfStock"
    End If

    productSchema = productSchema & Chr(34) & "," & _
        Chr(34) & "seller" & Chr(34) & ":{" & _
            Chr(34) & "@type" & Chr(34) & ":" & Chr(34) & "Organization" & Chr(34) & "," & _
            Chr(34) & "name" & Chr(34) & ":" & Chr(34) & "Özüm Klima" & Chr(34) & _
        "}}"
End If

productSchema = productSchema & "}"

breadcrumbSchema = "{" & _
    Chr(34) & "@context" & Chr(34) & ":" & Chr(34) & "https://schema.org" & Chr(34) & "," & _
    Chr(34) & "@type" & Chr(34) & ":" & Chr(34) & "BreadcrumbList" & Chr(34) & "," & _
    Chr(34) & "itemListElement" & Chr(34) & ":[" & _
        "{" & Chr(34) & "@type" & Chr(34) & ":" & Chr(34) & "ListItem" & Chr(34) & "," & Chr(34) & "position" & Chr(34) & ":1," & Chr(34) & "name" & Chr(34) & ":" & Chr(34) & "Ana Sayfa" & Chr(34) & "," & Chr(34) & "item" & Chr(34) & ":" & Chr(34) & "https://ozumklima.com/" & Chr(34) & "}," & _
        "{" & Chr(34) & "@type" & Chr(34) & ":" & Chr(34) & "ListItem" & Chr(34) & "," & Chr(34) & "position" & Chr(34) & ":2," & Chr(34) & "name" & Chr(34) & ":" & Chr(34) & JsonText(typeName) & Chr(34) & "," & Chr(34) & "item" & Chr(34) & ":" & Chr(34) & "https://ozumklima.com/klima.asp?t=" & CLng(uruns("tip")) & Chr(34) & "}," & _
        "{" & Chr(34) & "@type" & Chr(34) & ":" & Chr(34) & "ListItem" & Chr(34) & "," & Chr(34) & "position" & Chr(34) & ":3," & Chr(34) & "name" & Chr(34) & ":" & Chr(34) & JsonText(productName) & Chr(34) & "," & Chr(34) & "item" & Chr(34) & ":" & Chr(34) & JsonText(canonicalUrl) & Chr(34) & "}" & _
    "]}"

On Error Resume Next
baglanti.Execute "UPDATE products SET hit=IIF([hit] Is Null,1,[hit]+1) WHERE AffiliateID=" & CLng(pidi)
On Error GoTo 0
%>
<!doctype html>
<html lang="tr">
<head>
    <meta charset="utf-8">
    <meta name="viewport" content="width=device-width, initial-scale=1, viewport-fit=cover">
    <title><%=Server.HTMLEncode(seoTitle)%></title>
    <meta name="description" content="<%=Server.HTMLEncode(seoDescription)%>">
    <meta name="robots" content="index,follow,max-image-preview:large,max-snippet:-1,max-video-preview:-1">
    <link rel="canonical" href="<%=Server.HTMLEncode(canonicalUrl)%>">

    <meta property="og:locale" content="tr_TR">
    <meta property="og:type" content="product">
    <meta property="og:site_name" content="Özüm Klima">
    <meta property="og:title" content="<%=Server.HTMLEncode(seoTitle)%>">
    <meta property="og:description" content="<%=Server.HTMLEncode(seoDescription)%>">
    <meta property="og:url" content="<%=Server.HTMLEncode(canonicalUrl)%>">
<% If foto(1) <> "" Then %>
    <meta property="og:image" content="https://ozumklima.com/urunler/<%=Server.HTMLEncode(foto(1))%>">
    <meta property="og:image:alt" content="<%=Server.HTMLEncode(productName)%>">
<% End If %>
    <meta name="twitter:card" content="<% If foto(1) <> "" Then %>summary_large_image<% Else %>summary<% End If %>">
    <meta name="twitter:title" content="<%=Server.HTMLEncode(seoTitle)%>">
    <meta name="twitter:description" content="<%=Server.HTMLEncode(seoDescription)%>">
<% If effectivePrice > 0 Then %>
    <meta property="product:price:amount" content="<%=JsonNumber(effectivePrice)%>">
    <meta property="product:price:currency" content="TRY">
<% End If %>

    <link rel="stylesheet" href="css/font.css">
    <link rel="stylesheet" href="css/fontawesome-all.css">
    <link rel="stylesheet" href="css/darkbox.css">
    <link rel="shortcut icon" href="images/favicon.png">
    <script src="js/jquery-2.1.1.js"></script>

    <script type="application/ld+json"><%=productSchema%></script>
    <script type="application/ld+json"><%=breadcrumbSchema%></script>

    <style>
        .product-page,.product-page *{box-sizing:border-box}
        .product-page{width:calc(100% - 32px);max-width:1200px;margin:0 auto 64px;font-family:'Source Sans Pro',Arial,sans-serif;color:#333}
        .product-breadcrumb{padding:20px 0 12px;border-bottom:1px solid #ececec;font-size:.95rem;color:#666;overflow-wrap:anywhere}
        .product-breadcrumb a{color:#3e3e3e}
        .product-breadcrumb a:hover{text-decoration:underline}
        .product-breadcrumb span{margin:0 7px;color:#999}
        .product-hero{display:grid;grid-template-columns:minmax(0,1fr) minmax(340px,.9fr);gap:42px;align-items:start;padding:30px 0 36px}
        .product-gallery{min-width:0}
        .product-main-image{display:flex;width:100%;aspect-ratio:4/3;align-items:center;justify-content:center;overflow:hidden;background:#f7f7f7;border:1px solid #e4e4e4;border-radius:7px;cursor:zoom-in}
        .product-main-image img{display:block;width:100%;height:100%;object-fit:contain}
        .product-image-placeholder{display:flex;width:100%;aspect-ratio:4/3;align-items:center;justify-content:center;padding:28px;text-align:center;background:#f7f7f7;border:1px solid #e4e4e4;border-radius:7px;color:#777}
        .product-thumbs{display:grid;grid-template-columns:repeat(5,minmax(0,1fr));gap:9px;margin-top:10px}
        .product-thumb{padding:0;aspect-ratio:1;border:1px solid #ddd;border-radius:5px;background:#fff;overflow:hidden;cursor:pointer}
        .product-thumb img{display:block;width:100%;height:100%;object-fit:cover}
        .product-thumb.is-active{border:2px solid #8cac62}
        .product-summary{min-width:0}
        .product-summary h1{margin:0 0 15px;padding-bottom:14px;border-bottom:1px solid #ddd;font-size:clamp(1.8rem,3vw,2.55rem);line-height:1.12;letter-spacing:-.02em;color:#171717}
        .product-lead{margin:0 0 22px;font-size:1.08rem;line-height:1.65;color:#444}
        .product-price{margin:0 0 22px;font-size:1.85rem;font-weight:700;color:#202020}
        .product-old-price{display:block;margin-bottom:2px;font-size:.95rem;font-weight:400;text-decoration:line-through;color:#888}
        .product-price-note{font-size:1rem;font-weight:600;color:#5b5b5b}
        .product-facts{display:grid;grid-template-columns:minmax(120px,.45fr) 1fr;margin:0 0 23px;border-top:1px solid #e9e9e9}
        .product-facts dt,.product-facts dd{margin:0;padding:9px 0;border-bottom:1px solid #e9e9e9}
        .product-facts dt{font-weight:700;color:#555}
        .product-facts dd{overflow-wrap:anywhere;color:#222}
        .product-option-row{margin:0 0 18px}
        .product-option-row strong{display:block;margin-bottom:8px}
        .product-chips{display:flex;flex-wrap:wrap;gap:8px}
        .product-chip{display:inline-flex;align-items:center;gap:7px;min-height:34px;padding:6px 10px;border:1px solid #ddd;border-radius:5px;background:#fafafa;font-size:.92rem}
        .product-color-dot{width:18px;height:18px;border:1px solid rgba(0,0,0,.2);border-radius:50%}
        .product-resources{display:flex;flex-wrap:wrap;gap:10px;margin:21px 0}
        .product-resource{display:inline-flex;align-items:center;gap:9px;padding:10px 12px;border:1px solid #dedede;border-radius:5px;color:#333;background:#fff}
        .product-resource:hover{background:#f5f5f5}
        .product-resource img{width:32px;height:32px;object-fit:contain}
        .product-cta{display:inline-flex;width:auto;min-height:48px;align-items:center;justify-content:center;padding:12px 24px;border:0;border-radius:6px;background:#a3c225;color:#fff;font:700 1.05rem 'Source Sans Pro',Arial,sans-serif;cursor:pointer;transition:filter .2s ease,transform .2s ease}
        .product-cta:hover{filter:brightness(.93);transform:translateY(-1px)}
        .product-section{margin:0 0 34px;padding:30px;background:#fff;border:1px solid #e6e6e6;border-radius:7px}
        .product-section h2{margin:0 0 19px;font-size:1.65rem;color:#202020}
        .product-description-content{font-size:1.05rem;line-height:1.7;color:#333;overflow-wrap:anywhere}
        .product-description-content img,.product-description-content video{max-width:100%;height:auto}
        .product-description-content iframe{display:block;width:100%;max-width:100%;aspect-ratio:16/9;height:auto;border:0}
        .product-description-content table{display:block;width:100%;max-width:100%;overflow-x:auto;border-collapse:collapse}
        .product-video{margin-bottom:25px;overflow:hidden}
        .product-bottom-cta{text-align:center;margin-top:28px}
        .related-products{display:grid;grid-template-columns:repeat(4,minmax(0,1fr));gap:17px}
        .related-product{display:flex;min-width:0;flex-direction:column;border:1px solid #e4e4e4;border-radius:6px;background:#fff;overflow:hidden;transition:border-color .2s ease,transform .2s ease}
        .related-product:hover{border-color:#1f9bb7;transform:translateY(-2px)}
        .related-product__image{display:flex;aspect-ratio:4/3;align-items:center;justify-content:center;background:#f6f6f6;overflow:hidden}
        .related-product__image img{display:block;width:100%;height:100%;object-fit:contain}
        .related-product__placeholder{padding:18px;text-align:center;color:#888}
        .related-product__body{display:flex;flex:1;flex-direction:column;padding:14px}
        .related-product__name{margin:0 0 10px;font-size:1.05rem;font-weight:700;line-height:1.35;color:#111}
        .related-product__price{margin-top:auto;font-size:1.05rem;font-weight:700;color:#222}
        .related-product__old-price{display:block;font-size:.82rem;font-weight:400;text-decoration:line-through;color:#888}
        @media(max-width:920px){.product-hero{grid-template-columns:1fr;gap:26px}.product-gallery{max-width:720px;margin:0 auto}.related-products{grid-template-columns:repeat(2,minmax(0,1fr))}}
        @media(max-width:600px){.product-page{width:calc(100% - 20px);margin-bottom:40px}.product-breadcrumb{padding-top:14px;font-size:.86rem}.product-hero{padding-top:20px}.product-thumbs{grid-template-columns:repeat(4,minmax(0,1fr))}.product-summary h1{font-size:1.75rem}.product-facts{grid-template-columns:110px 1fr}.product-section{padding:21px 16px}.product-cta{width:100%}.related-products{grid-template-columns:1fr}.product-resources{flex-direction:column}.product-resource{width:100%}}
    </style>
</head>
<body>
<!--#include file='ust.asp'-->

<main class="product-page" id="main-content">
    <nav class="product-breadcrumb" aria-label="Sayfa yolu">
        <a href="/">Ana Sayfa</a>
        <span aria-hidden="true">/</span>
        <span><%=Server.HTMLEncode(categoryName)%></span>
        <span aria-hidden="true">/</span>
        <a href="klima.asp?t=<%=CLng(uruns("tip"))%>"><%=Server.HTMLEncode(typeName)%></a>
        <span aria-hidden="true">/</span>
        <strong><%=Server.HTMLEncode(productName)%></strong>
    </nav>

    <section class="product-hero" aria-labelledby="product-title">
        <div class="product-gallery">
<% If foto(1) <> "" Then %>
            <a id="productMainLink" class="product-main-image" href="urunler/<%=Server.HTMLEncode(foto(1))%>" data-darkbox="urunler/<%=Server.HTMLEncode(foto(1))%>" data-darkbox-description="<%=Server.HTMLEncode(productName)%>">
                <img id="productMainImage" src="urunler/<%=Server.HTMLEncode(foto(1))%>" width="720" height="540" alt="<%=Server.HTMLEncode(productName)%>" fetchpriority="high">
            </a>
<% Else %>
            <div class="product-image-placeholder" role="img" aria-label="<%=Server.HTMLEncode(productName)%> için henüz ürün görseli eklenmedi">
                Bu ürün için görsel yakında eklenecek.
            </div>
<% End If %>

<% If imageCount > 1 Then %>
            <div class="product-thumbs" aria-label="Ürün görselleri">
<% For imageIndex = 1 To 5 %>
<% If foto(imageIndex) <> "" Then %>
                <button type="button" class="product-thumb<% If imageIndex = 1 Then %> is-active<% End If %>" data-image="urunler/<%=Server.HTMLEncode(foto(imageIndex))%>" aria-label="<%=Server.HTMLEncode(productName)%> görseli <%=imageIndex%>">
                    <img src="urunler/<%=Server.HTMLEncode(foto(imageIndex))%>" width="130" height="130" alt="" loading="lazy">
                </button>
<% End If %>
<% Next %>
            </div>
<% End If %>
        </div>

        <div class="product-summary">
            <h1 id="product-title"><%=Server.HTMLEncode(productName)%></h1>
            <p class="product-lead"><%=Server.HTMLEncode(displayDescription)%></p>

<% If effectivePrice > 0 Then %>
            <div class="product-price" aria-label="Ürün fiyatı">
<% If discountPrice > 0 And regularPrice > discountPrice Then %>
                <span class="product-old-price"><%=FormatNumber(regularPrice, 2)%> TL</span>
<% End If %>
                <%=FormatNumber(effectivePrice, 2)%> TL
            </div>
<% Else %>
            <div class="product-price product-price-note">Fiyat ve uygulama koşulları için ücretsiz keşif isteyin.</div>
<% End If %>

            <dl class="product-facts">
<% If productCode <> "" And productCode <> "0" Then %>
                <dt>Ürün kodu</dt>
                <dd><%=Server.HTMLEncode(productCode)%></dd>
<% End If %>
                <dt>Kategori</dt>
                <dd><%=Server.HTMLEncode(typeName)%></dd>
<% If surfaceName <> "" Then %>
                <dt>Ürün grubu</dt>
                <dd><%=Server.HTMLEncode(surfaceName)%></dd>
<% End If %>
<% If deliveryDays > 0 Then %>
                <dt>Hazırlık süresi</dt>
                <dd>Yaklaşık <%=CLng(deliveryDays)%> gün</dd>
<% End If %>
                <dt>Hizmet</dt>
                <dd>Keşif, montaj ve teknik servis desteği</dd>
            </dl>

<% If hasColors Then %>
            <div class="product-option-row">
                <strong>Renk seçenekleri</strong>
                <div class="product-chips">
<% For optionIndex = 1 To 10 %>
<%
optionValue = Trim(uruns("renk" & optionIndex) & "")
If optionValue <> "" And optionValue <> "0" Then
%>
                    <span class="product-chip">
<% If IsHexColor(optionValue) Then %>
                        <span class="product-color-dot" style="background-color:#<%=Server.HTMLEncode(Replace(optionValue,"#",""))%>"></span>
<% End If %>
                        <%=Server.HTMLEncode(optionValue)%>
                    </span>
<% End If %>
<% Next %>
                </div>
            </div>
<% End If %>

<% If hasSizes Then %>
            <div class="product-option-row">
                <strong>Model / kapasite seçenekleri</strong>
                <div class="product-chips">
<% For optionIndex = 1 To 10 %>
<%
optionValue = Trim(uruns("beden" & optionIndex) & "")
If optionValue <> "" And optionValue <> "0" Then
%>
                    <span class="product-chip"><%=Server.HTMLEncode(optionValue)%></span>
<% End If %>
<% Next %>
                </div>
            </div>
<% End If %>

<% If Trim(uruns("uretici") & "") <> "" Or Trim(uruns("dokuman") & "") <> "" Then %>
            <div class="product-resources">
<% If Trim(uruns("uretici") & "") <> "" Then %>
                <a class="product-resource" href="urunler/<%=Server.HTMLEncode(uruns("uretici") & "")%>" target="_blank" rel="noopener">
                    <img src="images/ikon8.png" alt="" width="32" height="32">
                    <strong>Teknik özellikler</strong>
                </a>
<% End If %>
<% If Trim(uruns("dokuman") & "") <> "" Then %>
                <a class="product-resource" href="urunler/<%=Server.HTMLEncode(uruns("dokuman") & "")%>" target="_blank" rel="noopener">
                    <img src="images/ikon5.png" alt="" width="32" height="32">
                    <strong>Broşürü incele</strong>
                </a>
<% End If %>
            </div>
<% End If %>

            <form method="post" action="iletisim.asp" class="product-interest-form">
                <input type="hidden" name="fiyat" value="<%=effectivePrice%>">
                <input type="hidden" name="uid" value="<%=Server.HTMLEncode(Session("CustomerID") & "")%>">
                <input type="hidden" name="product" value="<%=CLng(uruns("AffiliateID"))%>">
                <button type="submit" class="product-cta" onclick="if(typeof analyticsTrack==='function'){analyticsTrack('product_interest_click',{cta:'ucretsiz_kesif',product_id:'<%=CLng(uruns("AffiliateID"))%>'});}">Ücretsiz keşif iste</button>
            </form>
        </div>
    </section>

<% If Trim(uruns("video") & "") <> "" Or Trim(uruns("notlar") & "") <> "" Then %>
    <section class="product-section" aria-labelledby="product-details-title">
        <h2 id="product-details-title">Ürün detayları</h2>
        <div class="product-description-content">
<% If Trim(uruns("video") & "") <> "" Then %>
            <div class="product-video"><%=uruns("video")%></div>
<% End If %>
            <%=uruns("notlar")%>
        </div>

        <div class="product-bottom-cta">
            <form method="post" action="iletisim.asp" class="product-interest-form">
                <input type="hidden" name="fiyat" value="<%=effectivePrice%>">
                <input type="hidden" name="uid" value="<%=Server.HTMLEncode(Session("CustomerID") & "")%>">
                <input type="hidden" name="product" value="<%=CLng(uruns("AffiliateID"))%>">
                <button type="submit" class="product-cta" onclick="if(typeof analyticsTrack==='function'){analyticsTrack('product_interest_click',{cta:'bu_urunle_ilgileniyorum',product_id:'<%=CLng(uruns("AffiliateID"))%>'});}">Bu ürünle ilgileniyorum</button>
            </form>
        </div>
    </section>
<% End If %>

<%
Dim relatedProducts, relatedPrice, relatedRegularPrice, relatedDiscountPrice, relatedDiscountPercent
Set relatedProducts = Server.CreateObject("ADODB.Recordset")
Sorgula = "SELECT TOP 8 products.* FROM related INNER JOIN products ON related.ilgili=products.AffiliateID " & _
          "WHERE related.urun=" & CLng(uruns("AffiliateID")) & " AND products.yayin=1 " & _
          "ORDER BY products.sira DESC, products.AffiliateID DESC"
relatedProducts.Open Sorgula, baglanti, 1, 1

If Not relatedProducts.EOF Then
%>
    <section class="product-section" aria-labelledby="related-products-title">
        <h2 id="related-products-title">Benzer ürünler</h2>
        <div class="related-products">
<%
Do While Not relatedProducts.EOF
    relatedRegularPrice = NumberOrZero(relatedProducts("fiyat"))
    relatedDiscountPrice = NumberOrZero(relatedProducts("indirim"))
    relatedPrice = relatedRegularPrice
    If relatedDiscountPrice > 0 Then relatedPrice = relatedDiscountPrice
    relatedDiscountPercent = 0

    If relatedRegularPrice > 0 Then
        If relatedDiscountPrice > 0 Then
            If relatedDiscountPrice < relatedRegularPrice Then
                relatedDiscountPercent = Round((1 - (relatedDiscountPrice / relatedRegularPrice)) * 100)
            End If
        End If
    End If
%>
            <a class="related-product" href="<%=ProductSlug(relatedProducts("isim") & "")%>-<%=CLng(relatedProducts("AffiliateID"))%>">
                <div class="related-product__image">
<% If Trim(relatedProducts("foto1") & "") <> "" Then %>
                    <img src="urunler/<%=Server.HTMLEncode(relatedProducts("foto1") & "")%>" width="360" height="270" loading="lazy" alt="<%=Server.HTMLEncode(relatedProducts("isim") & "")%>">
<% Else %>
                    <span class="related-product__placeholder">Görsel yakında eklenecek</span>
<% End If %>
                </div>
                <div class="related-product__body">
                    <div class="related-product__name"><%=Server.HTMLEncode(relatedProducts("isim") & "")%></div>
<% If relatedPrice > 0 Then %>
                    <div class="related-product__price">
<% If relatedDiscountPercent > 0 Then %>
                        <span class="related-product__old-price"><%=FormatNumber(relatedRegularPrice,2)%> TL · %<%=relatedDiscountPercent%> indirim</span>
<% End If %>
                        <%=FormatNumber(relatedPrice,2)%> TL
                    </div>
<% End If %>
                </div>
            </a>
<%
    relatedProducts.MoveNext
Loop
%>
        </div>
    </section>
<% End If %>
<%
If relatedProducts.State = 1 Then relatedProducts.Close
Set relatedProducts = Nothing
%>
</main>

<!--#include file='alt.asp'-->
<script src="js/darkbox.js"></script>
<script>
$(function(){
    $('.product-thumb').on('click',function(){
        var button=$(this);
        var imagePath=button.data('image');

        $('.product-thumb').removeClass('is-active');
        button.addClass('is-active');
        $('#productMainImage').attr('src',imagePath);
        $('#productMainLink').attr('href',imagePath).attr('data-darkbox',imagePath);
    });
});
</script>
<%
If uruns.State = 1 Then uruns.Close
Set uruns = Nothing
If anas.State = 1 Then anas.Close
Set anas = Nothing
If tipi.State = 1 Then tipi.Close
Set tipi = Nothing
If Not (sorf Is Nothing) Then
    If sorf.State = 1 Then sorf.Close
    Set sorf = Nothing
End If
%>
</body>
</html>
