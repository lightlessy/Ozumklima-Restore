<!doctype html>
<html lang="tr">
<head>
    <meta charset="utf-8">
    <meta name="viewport" content="width=device-width, initial-scale=1, viewport-fit=cover">
<% If productFound Then %>
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
    <meta name="twitter:card" content="summary_large_image">
<% Else %>
    <meta name="twitter:card" content="summary">
<% End If %>
    <meta name="twitter:title" content="<%=Server.HTMLEncode(seoTitle)%>">
    <meta name="twitter:description" content="<%=Server.HTMLEncode(seoDescription)%>">
<% If effectivePrice > 0 Then %>
    <meta property="product:price:amount" content="<%=DpJsonNumber(effectivePrice)%>">
    <meta property="product:price:currency" content="TRY">
<% End If %>
<% ElseIf dpFatalError Then %>
    <title>Sayfa geçici olarak kullanılamıyor | Özüm Klima</title>
    <meta name="robots" content="noindex,nofollow">
<% Else %>
    <title>Ürün bulunamadı | Özüm Klima</title>
    <meta name="robots" content="noindex,follow">
<% End If %>

    <link rel="stylesheet" href="css/font.css">
    <link rel="stylesheet" href="css/fontawesome-all.css">
    <link rel="stylesheet" href="css/darkbox.css">
    <link rel="shortcut icon" href="images/favicon.png">
    <script src="js/jquery-2.1.1.js"></script>

<% If productFound Then %>
    <script type="application/ld+json"><%=productSchema%></script>
    <script type="application/ld+json"><%=breadcrumbSchema%></script>
<% End If %>

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
        .product-not-found{width:calc(100% - 32px);max-width:900px;margin:70px auto;padding:45px 24px;box-sizing:border-box;text-align:center;background:#fff;border:1px solid #e2e2e2;font-family:'Source Sans Pro',Arial,sans-serif;color:#333}
        .product-not-found h1{margin:0 0 12px;font-size:2rem}
        .product-not-found a{display:inline-block;margin-top:16px;padding:11px 18px;background:#3c3c3c;color:#fff;border-radius:5px}
    </style>
</head>
<body>
