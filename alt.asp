<%
Response.CodePage = 65001
Response.Charset = "utf-8"

Function FooterSafe(ByVal value)
  FooterSafe = Server.HTMLEncode(CStr(value & ""))
End Function

Function NormalizeWhatsAppNumber(ByVal value)
  Dim numberValue

  numberValue = CStr(value & "")
  numberValue = Replace(numberValue, " ", "")
  numberValue = Replace(numberValue, "(", "")
  numberValue = Replace(numberValue, ")", "")
  numberValue = Replace(numberValue, "-", "")
  numberValue = Replace(numberValue, "+", "")

  If Left(numberValue, 2) = "90" Then
    NormalizeWhatsAppNumber = numberValue
  ElseIf Left(numberValue, 1) = "0" Then
    NormalizeWhatsAppNumber = "9" & numberValue
  Else
    NormalizeWhatsAppNumber = "90" & numberValue
  End If
End Function

Sub RenderFooterProducts(ByVal productType, ByVal limitCount)
  Dim footerProducts
  Dim footerSql

  Set footerProducts = Server.CreateObject("ADODB.Recordset")
  footerSql = "SELECT TOP " & CInt(limitCount) & _
              " * FROM products WHERE tip = " & CInt(productType) & _
              " ORDER BY sira DESC"

  footerProducts.Open footerSql, baglanti, 1, 3

  Do While Not footerProducts.Eof
%>
    <a href="<%=cevir(footerProducts("isim"))%>-<%=footerProducts("AffiliateID")%>">
      <%=FooterSafe(Left(footerProducts("isim"), 30))%>
    </a>
<%
    footerProducts.MoveNext
  Loop

  footerProducts.Close
  Set footerProducts = Nothing
End Sub

Dim whatsappNumber
Dim whatsappMessage
Dim whatsappUrl

whatsappNumber = NormalizeWhatsAppNumber(ayars("gsm"))
whatsappMessage = Server.URLEncode("Merhaba, Özüm Klima hizmetleri hakkında bilgi almak istiyorum.")
whatsappUrl = "https://wa.me/" & whatsappNumber & "?text=" & whatsappMessage
%>

<style>
  .whatsappCta {
    position: fixed;
    right: max(18px, env(safe-area-inset-right));
    bottom: max(18px, env(safe-area-inset-bottom));
    z-index: 9999;
    font-family: Arial, Helvetica, sans-serif;
  }

  .whatsappCta a {
    display: inline-flex;
    align-items: center;
    gap: 11px;
    min-height: 52px;
    padding: 9px 16px 9px 10px;
    border: 1px solid rgba(255, 255, 255, 0.34);
    border-radius: 999px;
    background: rgba(255, 255, 255, 0.72);
    color: #173E2A;
    text-decoration: none;
    box-shadow:
      0 14px 36px rgba(0, 0, 0, 0.16),
      inset 0 1px 0 rgba(255, 255, 255, 0.55);
    backdrop-filter: blur(16px) saturate(145%);
    -webkit-backdrop-filter: blur(16px) saturate(145%);
    transition:
      transform 0.2s ease,
      box-shadow 0.2s ease,
      background 0.2s ease,
      border-color 0.2s ease;
  }

  .whatsappCta a:hover {
    transform: translateY(-2px);
    background: rgba(255, 255, 255, 0.9);
    border-color: rgba(37, 211, 102, 0.42);
    box-shadow:
      0 18px 42px rgba(0, 0, 0, 0.19),
      0 0 0 5px rgba(37, 211, 102, 0.08);
  }

  .whatsappCta a:focus-visible {
    outline: 3px solid rgba(37, 211, 102, 0.65);
    outline-offset: 4px;
  }

  .whatsappCtaIcon {
    display: inline-flex;
    align-items: center;
    justify-content: center;
    width: 38px;
    height: 38px;
    flex: 0 0 38px;
    border-radius: 50%;
    background: #25D366;
    color: #FFFFFF;
    box-shadow: 0 8px 20px rgba(37, 211, 102, 0.28);
  }

  .whatsappCtaIcon svg {
    width: 21px;
    height: 21px;
    display: block;
  }

  .whatsappCtaText {
    display: flex;
    flex-direction: column;
    line-height: 1.15;
  }

  .whatsappCtaText strong {
    color: #173E2A;
    font-size: 0.88rem;
    font-weight: 700;
  }

  .whatsappCtaText span {
    margin-top: 3px;
    color: #5F6F65;
    font-size: 0.72rem;
    font-weight: 600;
  }

  .siteFooter {
    background: #353535;
    color: #F1F1F1;
    font-family: Arial, Helvetica, sans-serif;
  }

  .siteFooterInner {
    width: 92%;
    max-width: 1180px;
    margin: 0 auto;
    padding: 54px 0 42px;
  }

  .siteFooterGrid {
    display: grid;
    grid-template-columns: minmax(0, 1fr) 220px;
    gap: 48px;
    align-items: start;
  }

  .siteFooterMenus {
    display: grid;
    grid-template-columns: repeat(3, minmax(150px, 1fr));
    gap: 34px 28px;
  }

  .siteFooterColumn h3 {
    margin: 0 0 14px;
    color: #FFFFFF;
    font-size: 0.98rem;
    line-height: 1.3;
  }

  .siteFooterColumn h3 a {
    color: inherit;
    text-decoration: none;
  }

  .siteFooterColumn a {
    display: block;
    margin: 0 0 8px;
    color: #CFCFCF;
    font-size: 0.83rem;
    line-height: 1.45;
    text-decoration: none;
    transition: color 0.18s ease;
  }

  .siteFooterColumn a:hover {
    color: #FFFFFF;
  }

  .siteFooterContact h3 {
    margin: 0 0 16px;
    color: #FFFFFF;
    font-size: 1rem;
  }

  .siteFooterContact p {
    margin: 0 0 14px;
    color: #D7D7D7;
    font-size: 0.86rem;
    line-height: 1.6;
  }

  .siteFooterContact a {
    color: #DBF2EB;
    text-decoration: none;
  }

  .siteFooterContact a:hover {
    text-decoration: underline;
  }

  .siteFooterPhone {
    display: inline-flex;
    align-items: center;
    gap: 8px;
    margin-bottom: 14px;
    color: #FFFFFF;
    font-weight: 700;
  }

  .siteFooterBottom {
    border-top: 1px solid #555555;
    background: #353535;
  }

  .siteFooterBottomInner {
    width: 92%;
    max-width: 1180px;
    margin: 0 auto;
    padding: 18px 0;
    color: #BEBEBE;
    font-family: Arial, Helvetica, sans-serif;
    font-size: 0.78rem;
    line-height: 1.5;
  }

  .siteFooterBottomInner a {
    color: #E6E6E6;
    text-decoration: none;
  }

  .siteFooterBottomInner a:hover {
    color: #FFFFFF;
    text-decoration: underline;
  }

  @media only screen and (max-width: 900px) {
    .siteFooterGrid {
      grid-template-columns: 1fr;
      gap: 38px;
    }

    .siteFooterMenus {
      grid-template-columns: repeat(2, minmax(0, 1fr));
    }

    .siteFooterContact {
      max-width: 420px;
    }
  }

  @media only screen and (max-width: 600px) {
    .whatsappCta {
      right: max(12px, env(safe-area-inset-right));
      bottom: max(12px, env(safe-area-inset-bottom));
    }

    .whatsappCta a {
      width: 52px;
      height: 52px;
      min-height: 52px;
      padding: 0;
      justify-content: center;
      border-radius: 50%;
    }

    .whatsappCtaIcon {
      width: 40px;
      height: 40px;
      flex-basis: 40px;
    }

    .whatsappCtaText {
      position: absolute;
      width: 1px;
      height: 1px;
      overflow: hidden;
      clip: rect(0 0 0 0);
      white-space: nowrap;
    }

    .siteFooterInner {
      width: calc(100% - 28px);
      padding: 40px 0 34px;
    }

    .siteFooterMenus {
      grid-template-columns: 1fr;
      gap: 26px;
    }

    .siteFooterBottomInner {
      width: calc(100% - 28px);
      text-align: center;
    }
  }

  @media (prefers-reduced-motion: reduce) {
    .whatsappCta a,
    .siteFooterColumn a {
      transition: none;
    }
  }
</style>

<div class="whatsappCta">
  <a
    href="<%=whatsappUrl%>"
    target="_blank"
    rel="noopener noreferrer"
    aria-label="WhatsApp üzerinden Özüm Klima destek hattına ulaşın"
    onclick="analyticsTrack('whatsapp_click', {location: 'footer', method: 'wa.me'});"
  >
    <span class="whatsappCtaIcon" aria-hidden="true">
      <svg viewBox="0 0 32 32" role="img" focusable="false">
        <path
          fill="currentColor"
          d="M16.04 3.2c-6.98 0-12.66 5.67-12.66 12.65 0 2.23.58 4.41 1.69 6.33L3.27 28.8l6.78-1.78a12.63 12.63 0 0 0 5.98 1.52h.01c6.98 0 12.66-5.68 12.66-12.66 0-3.38-1.32-6.56-3.71-8.95A12.57 12.57 0 0 0 16.04 3.2Zm0 23.2h-.01a10.48 10.48 0 0 1-5.34-1.46l-.38-.23-4.02 1.06 1.07-3.92-.25-.4a10.51 10.51 0 1 1 8.93 4.95Zm5.77-7.88c-.32-.16-1.87-.92-2.16-1.03-.29-.1-.5-.16-.71.16-.21.32-.81 1.03-1 1.24-.18.21-.37.24-.68.08-.32-.16-1.34-.49-2.55-1.58-.94-.84-1.58-1.88-1.76-2.2-.19-.31-.02-.48.14-.64.14-.14.31-.37.47-.55.16-.19.21-.32.32-.53.1-.21.05-.4-.03-.55-.08-.16-.71-1.72-.98-2.35-.26-.62-.52-.54-.71-.55h-.61c-.21 0-.55.08-.84.4-.29.32-1.11 1.08-1.11 2.64 0 1.56 1.13 3.06 1.29 3.27.16.21 2.23 3.4 5.4 4.77.75.33 1.34.52 1.8.66.76.24 1.44.21 1.99.13.61-.09 1.87-.77 2.14-1.5.26-.74.26-1.37.18-1.5-.08-.13-.29-.21-.61-.37Z"
        />
      </svg>
    </span>

    <span class="whatsappCtaText">
      <strong>WhatsApp Destek</strong>
      <span>Hızlıca bize ulaşın</span>
    </span>
  </a>
</div>

<footer class="siteFooter">
  <div class="siteFooterInner">
    <div class="siteFooterGrid">

      <nav class="siteFooterMenus" aria-label="Alt menü">

        <section class="siteFooterColumn">
          <h3><a href="klima.asp?t=36">Ev ve Ofis Klimaları</a></h3>
          <% Call RenderFooterProducts(36, 9) %>
        </section>

        <section class="siteFooterColumn">
          <h3><a href="klima.asp?t=37">Ticari Tip Klimalar</a></h3>
          <% Call RenderFooterProducts(37, 9) %>
        </section>

        <section class="siteFooterColumn">
          <h3><a href="vrf_sistemler.asp">VRF Merkezi Sistemler</a></h3>
          <% Call RenderFooterProducts(41, 9) %>
        </section>

        <section class="siteFooterColumn">
          <h3><a href="havalandirma.asp">Havalandırma Sistemleri</a></h3>
          <% Call RenderFooterProducts(39, 9) %>
        </section>

        <section class="siteFooterColumn">
          <h3><a href="h.asp?id=12">Antalya Hizmetleri</a></h3>
          <a href="antalya-klima-servisi.asp">Antalya Klima Servisi</a>
          <a href="dosemealti-klima-servisi.asp">Döşemealtı Klima Servisi</a>
          <a href="antalya-klima-montaji.asp">Antalya Klima Montajı</a>
          <a href="antalya-klima-bakimi.asp">Antalya Klima Bakımı</a>
          <a href="antalya-mitsubishi-electric-klima.asp">Mitsubishi Electric Klima</a>
          <a href="antalya-vrf-sistemleri.asp">Antalya VRF Sistemleri</a>
        </section>

        <section class="siteFooterColumn">
          <h3><a href="klima-rehberi.asp">Klima Rehberi</a></h3>
          <a href="40-m2-oda-kac-btu-klima.asp">40 m² oda için kaç BTU klima gerekir?</a>
          <a href="vrf-mi-multi-split-mi.asp">VRF mi, Multi Split mi?</a>
          <a href="klima-kesfi-nasil-yapilir.asp">Klima keşfi nasıl yapılır?</a>
          <a href="antalya-villa-klima-cozumleri.asp">Antalya villa klima çözümleri</a>
        </section>

      </nav>

      <aside class="siteFooterContact" aria-labelledby="footer-contact-title">
        <h3 id="footer-contact-title">Bize Ulaşın</h3>

        <div class="siteFooterPhone">
          <i class="fas fa-phone" aria-hidden="true"></i>
          <span><%=FooterSafe(ayars("tel"))%></span>
        </div>

        <p>
          <strong>Adres</strong><br>
          <%=FooterSafe(ayars("adres"))%>
        </p>

        <p>
          <strong>E-posta</strong><br>
          <a href="mailto:<%=FooterSafe(ayars("eposta"))%>">
            <%=FooterSafe(ayars("eposta"))%>
          </a>
        </p>
      </aside>

    </div>
  </div>
</footer>

<div class="siteFooterBottom">
  <div class="siteFooterBottomInner">
    &copy; <%=Year(Date)%>
    <a href="https://www.ozumklima.com" title="Antalya klima çözümleri">Özüm Klima</a>.
    Tüm hakları saklıdır.
    &nbsp;|&nbsp;
    <a href="https://www.megatasarim.com" target="_blank" rel="noopener noreferrer">
      Mega Tasarım
    </a>
  </div>
</div>

<!--#include file='analytics.asp'-->
