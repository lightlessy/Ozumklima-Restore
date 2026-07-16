
<% If dpFatalError Then %>
<main class="product-not-found" id="main-content">
    <h1>Sayfa geçici olarak kullanılamıyor</h1>
    <p>Ürün bilgileri şu anda yüklenemedi. Lütfen kısa süre sonra tekrar deneyin.</p>
    <a href="/">Ana sayfaya dön</a>
</main>
<% ElseIf Not productFound Then %>
<main class="product-not-found" id="main-content">
    <h1>Ürün bulunamadı</h1>
    <p>Aradığınız ürün kaldırılmış, yayından alınmış veya bağlantı hatalı olabilir.</p>
    <a href="/">Ana sayfaya dön</a>
</main>
<% Else %>
<main class="product-page" id="main-content">
    <nav class="product-breadcrumb" aria-label="Sayfa yolu">
        <a href="/">Ana Sayfa</a>
        <span aria-hidden="true">/</span>
        <span><%=Server.HTMLEncode(categoryName)%></span>
        <span aria-hidden="true">/</span>
        <a href="klima.asp?t=<%=productTypeId%>"><%=Server.HTMLEncode(typeName)%></a>
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
<% optionValue = colorValues(optionIndex) %>
<% If optionValue <> "" And optionValue <> "0" Then %>
                    <span class="product-chip">
<% If DpIsHexColor(optionValue) Then %>
<% cleanColor = Replace(optionValue, "#", "") %>
                        <span class="product-color-dot" data-color="#<%=Server.HTMLEncode(cleanColor)%>"></span>
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
<% optionValue = sizeValues(optionIndex) %>
<% If optionValue <> "" And optionValue <> "0" Then %>
                    <span class="product-chip"><%=Server.HTMLEncode(optionValue)%></span>
<% End If %>
<% Next %>
                </div>
            </div>
<% End If %>

<% If manufacturerFile <> "" Or documentFile <> "" Then %>
            <div class="product-resources">
<% If manufacturerFile <> "" Then %>
                <a class="product-resource" href="urunler/<%=Server.HTMLEncode(manufacturerFile)%>" target="_blank" rel="noopener">
                    <img src="images/ikon8.png" alt="" width="32" height="32">
                    <strong>Teknik özellikler</strong>
                </a>
<% End If %>
<% If documentFile <> "" Then %>
                <a class="product-resource" href="urunler/<%=Server.HTMLEncode(documentFile)%>" target="_blank" rel="noopener">
                    <img src="images/ikon5.png" alt="" width="32" height="32">
                    <strong>Broşürü incele</strong>
                </a>
<% End If %>
            </div>
<% End If %>

            <form method="post" action="iletisim.asp" class="product-interest-form">
                <input type="hidden" name="fiyat" value="<%=effectivePrice%>">
                <input type="hidden" name="uid" value="<%=Server.HTMLEncode(Session("CustomerID") & "")%>">
                <input type="hidden" name="product" value="<%=productId%>">
                <button type="submit" class="product-cta js-product-interest" data-cta="ucretsiz_kesif" data-product-id="<%=productId%>">Ücretsiz keşif iste</button>
            </form>
        </div>
    </section>

<% If Trim(productVideo) <> "" Or Trim(productNotes) <> "" Then %>
    <section class="product-section" aria-labelledby="product-details-title">
        <h2 id="product-details-title">Ürün detayları</h2>
        <div class="product-description-content">
<% If Trim(productVideo) <> "" Then %>
            <div class="product-video"><%=productVideo%></div>
<% End If %>
            <%=productNotes%>
        </div>

        <div class="product-bottom-cta">
            <form method="post" action="iletisim.asp" class="product-interest-form">
                <input type="hidden" name="fiyat" value="<%=effectivePrice%>">
                <input type="hidden" name="uid" value="<%=Server.HTMLEncode(Session("CustomerID") & "")%>">
                <input type="hidden" name="product" value="<%=productId%>">
                <button type="submit" class="product-cta js-product-interest" data-cta="bu_urunle_ilgileniyorum" data-product-id="<%=productId%>">Bu ürünle ilgileniyorum</button>
            </form>
        </div>
    </section>
<% End If %>

<% If hasRelatedProducts Then %>
    <section class="product-section" aria-labelledby="related-products-title">
        <h2 id="related-products-title">Benzer ürünler</h2>
        <div class="related-products">
<%
Do While Not relatedProducts.EOF
    relatedName = Trim(DpText(DpField(relatedProducts, "isim", "")))
    relatedId = DpPositiveLong(DpField(relatedProducts, "AffiliateID", 0))
    relatedImage = Trim(DpText(DpField(relatedProducts, "foto1", "")))
    relatedUrl = DpProductSlug(relatedName) & "-" & relatedId
    relatedRegularPrice = DpNumber(DpField(relatedProducts, "fiyat", 0))
    relatedDiscountPrice = DpNumber(DpField(relatedProducts, "indirim", 0))
    relatedPrice = relatedRegularPrice
    If relatedDiscountPrice > 0 Then relatedPrice = relatedDiscountPrice
    relatedDiscountPercent = 0

    If relatedRegularPrice > 0 And relatedDiscountPrice > 0 Then
        If relatedDiscountPrice < relatedRegularPrice Then
            relatedDiscountPercent = Round((1 - (relatedDiscountPrice / relatedRegularPrice)) * 100)
        End If
    End If
%>
            <a class="related-product" href="<%=Server.HTMLEncode(relatedUrl)%>">
                <div class="related-product__image">
<% If relatedImage <> "" Then %>
                    <img src="urunler/<%=Server.HTMLEncode(relatedImage)%>" width="360" height="270" loading="lazy" alt="<%=Server.HTMLEncode(relatedName)%>">
<% Else %>
                    <span class="related-product__placeholder">Görsel yakında eklenecek</span>
<% End If %>
                </div>
                <div class="related-product__body">
                    <div class="related-product__name"><%=Server.HTMLEncode(relatedName)%></div>
<% If relatedPrice > 0 Then %>
                    <div class="related-product__price">
<% If relatedDiscountPercent > 0 Then %>
                        <span class="related-product__old-price"><%=FormatNumber(relatedRegularPrice, 2)%> TL · %<%=relatedDiscountPercent%> indirim</span>
<% End If %>
                        <%=FormatNumber(relatedPrice, 2)%> TL
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
</main>
<% End If %>

