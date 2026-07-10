<%
Function UygulamaCaseSql(ByVal value)
  UygulamaCaseSql = Replace(CStr(value & ""), "'", "''")
End Function

Function UygulamaCaseText(ByVal value)
  Dim textValue
  textValue = Server.HTMLEncode(CStr(value & ""))
  textValue = Replace(textValue, vbCrLf, "<br>")
  textValue = Replace(textValue, vbLf, "<br>")
  UygulamaCaseText = textValue
End Function

Function UygulamaCaseAttr(ByVal value)
  UygulamaCaseAttr = Server.HTMLEncode(CStr(value & ""))
End Function

Function UygulamaCaseJs(ByVal value)
  Dim jsValue
  jsValue = CStr(value & "")
  jsValue = Replace(jsValue, "\", "\\")
  jsValue = Replace(jsValue, "'", "\'")
  jsValue = Replace(jsValue, Chr(34), "&quot;")
  jsValue = Replace(jsValue, vbCrLf, "")
  jsValue = Replace(jsValue, vbLf, "")
  UygulamaCaseJs = jsValue
End Function

Function UygulamaCaseValue(ByRef rs, ByVal fieldName)
  On Error Resume Next
  UygulamaCaseValue = rs(fieldName) & ""
  If Err.Number <> 0 Then
    UygulamaCaseValue = ""
    Err.Clear
  End If
  On Error GoTo 0
End Function

Function UygulamaCaseUrl(ByVal value)
  Dim urlValue, lowerUrl
  urlValue = Trim(CStr(value & ""))
  If urlValue = "" Then urlValue = "iletisim.asp"
  If Left(urlValue, 3) = "../" Then urlValue = Mid(urlValue, 4)

  lowerUrl = LCase(urlValue)
  If InStr(1, lowerUrl, "javascript:", 1) > 0 Then urlValue = "iletisim.asp"

  UygulamaCaseUrl = Server.HTMLEncode(urlValue)
End Function

Function UygulamaCaseImage(ByVal value)
  Dim imageValue, lowerImage
  imageValue = Trim(CStr(value & ""))
  lowerImage = LCase(imageValue)
  If InStr(1, lowerImage, "javascript:", 1) > 0 Then imageValue = ""
  UygulamaCaseImage = Server.HTMLEncode(imageValue)
End Function

Function UygulamaCaseDetailUrl(ByVal slug)
  Dim slugValue
  slugValue = Trim(CStr(slug & ""))
  If slugValue = "" Then slugValue = "uygulama-ornegi"
  UygulamaCaseDetailUrl = "uygulama-ornegi.asp?slug=" & Server.URLEncode(slugValue)
End Function

Sub UygulamaCaseAddColumn(ByVal columnSql)
  On Error Resume Next
  baglanti.Execute "ALTER TABLE uygulama_ornekleri ADD COLUMN " & columnSql
  Err.Clear
  On Error GoTo 0
End Sub

Sub UygulamaCaseSeed(ByVal baslik, ByVal etiket, ByVal kisaOzet, ByVal lokasyon, ByVal projeTipi, ByVal hizmetTipi, ByVal sistemMarka, ByVal projeSuresi, ByVal problem, ByVal kesif, ByVal cozum, ByVal uygulama, ByVal sonuc, ByVal ctaMetni, ByVal ctaUrl, ByVal slug, ByVal takipEtiketi, ByVal sira)
  Dim seedSql
  seedSql = "INSERT INTO uygulama_ornekleri ([baslik], [etiket], [kisa_ozet], [lokasyon], [proje_tipi], [hizmet_tipi], [sistem_marka], [proje_suresi], [problem], [kesif_degerlendirme], [cozum], [uygulama], [sonuc], [cta_metni], [cta_url], [slug], [takip_etiketi], [sira], [yayin], [kayit_tarihi]) VALUES (" & _
            "'" & UygulamaCaseSql(baslik) & "'," & _
            "'" & UygulamaCaseSql(etiket) & "'," & _
            "'" & UygulamaCaseSql(kisaOzet) & "'," & _
            "'" & UygulamaCaseSql(lokasyon) & "'," & _
            "'" & UygulamaCaseSql(projeTipi) & "'," & _
            "'" & UygulamaCaseSql(hizmetTipi) & "'," & _
            "'" & UygulamaCaseSql(sistemMarka) & "'," & _
            "'" & UygulamaCaseSql(projeSuresi) & "'," & _
            "'" & UygulamaCaseSql(problem) & "'," & _
            "'" & UygulamaCaseSql(kesif) & "'," & _
            "'" & UygulamaCaseSql(cozum) & "'," & _
            "'" & UygulamaCaseSql(uygulama) & "'," & _
            "'" & UygulamaCaseSql(sonuc) & "'," & _
            "'" & UygulamaCaseSql(ctaMetni) & "'," & _
            "'" & UygulamaCaseSql(ctaUrl) & "'," & _
            "'" & UygulamaCaseSql(slug) & "'," & _
            "'" & UygulamaCaseSql(takipEtiketi) & "'," & CLng(sira) & ",-1,Now())"
  baglanti.Execute seedSql
End Sub

Sub UygulamaCaseEnsureTable()
  Dim testRs, toplamKayit
  toplamKayit = 0

  On Error Resume Next
  Set testRs = baglanti.Execute("SELECT TOP 1 [id] FROM uygulama_ornekleri")
  If Err.Number <> 0 Then
    Err.Clear
    baglanti.Execute "CREATE TABLE uygulama_ornekleri ([id] AUTOINCREMENT PRIMARY KEY, [baslik] TEXT(255), [etiket] TEXT(100), [kisa_ozet] MEMO, [lokasyon] TEXT(150), [proje_tipi] TEXT(120), [hizmet_tipi] TEXT(150), [sistem_marka] TEXT(180), [proje_suresi] TEXT(80), [problem] MEMO, [kesif_degerlendirme] MEMO, [cozum] MEMO, [uygulama] MEMO, [sonuc] MEMO, [cover_image] TEXT(255), [cover_alt] TEXT(255), [galeri_gorseller] MEMO, [galeri_altlar] MEMO, [video_url] TEXT(255), [oncesi_image] TEXT(255), [oncesi_alt] TEXT(255), [sonrasi_image] TEXT(255), [sonrasi_alt] TEXT(255), [musteri_yorumu] MEMO, [sss] MEMO, [ilgili_hizmet_url] TEXT(255), [ilgili_bolge_url] TEXT(255), [cta_metni] TEXT(255), [cta_url] TEXT(255), [slug] TEXT(120), [takip_etiketi] TEXT(120), [seo_title] TEXT(255), [meta_description] MEMO, [sira] INTEGER, [yayin] YESNO, [kayit_tarihi] DATETIME)"
  End If
  If Not testRs Is Nothing Then
    If testRs.State = 1 Then testRs.Close
    Set testRs = Nothing
  End If
  On Error GoTo 0

  Call UygulamaCaseAddColumn("[kisa_ozet] MEMO")
  Call UygulamaCaseAddColumn("[lokasyon] TEXT(150)")
  Call UygulamaCaseAddColumn("[proje_tipi] TEXT(120)")
  Call UygulamaCaseAddColumn("[hizmet_tipi] TEXT(150)")
  Call UygulamaCaseAddColumn("[sistem_marka] TEXT(180)")
  Call UygulamaCaseAddColumn("[proje_suresi] TEXT(80)")
  Call UygulamaCaseAddColumn("[kesif_degerlendirme] MEMO")
  Call UygulamaCaseAddColumn("[cover_image] TEXT(255)")
  Call UygulamaCaseAddColumn("[cover_alt] TEXT(255)")
  Call UygulamaCaseAddColumn("[galeri_gorseller] MEMO")
  Call UygulamaCaseAddColumn("[galeri_altlar] MEMO")
  Call UygulamaCaseAddColumn("[video_url] TEXT(255)")
  Call UygulamaCaseAddColumn("[oncesi_image] TEXT(255)")
  Call UygulamaCaseAddColumn("[oncesi_alt] TEXT(255)")
  Call UygulamaCaseAddColumn("[sonrasi_image] TEXT(255)")
  Call UygulamaCaseAddColumn("[sonrasi_alt] TEXT(255)")
  Call UygulamaCaseAddColumn("[musteri_yorumu] MEMO")
  Call UygulamaCaseAddColumn("[sss] MEMO")
  Call UygulamaCaseAddColumn("[ilgili_hizmet_url] TEXT(255)")
  Call UygulamaCaseAddColumn("[ilgili_bolge_url] TEXT(255)")
  Call UygulamaCaseAddColumn("[takip_etiketi] TEXT(120)")
  Call UygulamaCaseAddColumn("[seo_title] TEXT(255)")
  Call UygulamaCaseAddColumn("[meta_description] MEMO")

  On Error Resume Next
  Set testRs = baglanti.Execute("SELECT COUNT(*) AS toplam FROM uygulama_ornekleri")
  If Err.Number = 0 Then toplamKayit = CLng(testRs("toplam"))
  If Not testRs Is Nothing Then
    If testRs.State = 1 Then testRs.Close
    Set testRs = Nothing
  End If
  On Error GoTo 0

  If toplamKayit = 0 Then
    Call UygulamaCaseSeed("Antalya Villa Klima Montaji", "Villa / Konut", "Konyaalti bolgesindeki villada salon ve yatak odalari icin sessiz, dengeli ve estetik klima cozumunun nasil planlandigini inceleyin.", "Konyaalti / Antalya", "Villa", "Klima montaji", "Mitsubishi Electric multi split", "1 gun", "Musteri yaz aylarinda salon ve yatak odalarinda yeterli sogutma alamiyor, dis unite konumunun goruntu ve ses acisindan sorun yaratmamasini istiyordu.", "Kesifte oda metrekareleri, cephe yonu, cam alani, drenaj hattinin gececegi guzergah, elektrik altyapisi ve dis unite servis erisimi kontrol edildi.", "Alan ihtiyacina uygun kapasite hesaplandi ve birden fazla ic uniteyi dengeli calistirabilecek Mitsubishi Electric cozum onerildi.", "Ic unite yerleri musteriyle netlestirildi, boru ve drenaj hatlari gizli guzergahlarla tasindi, dis unite bakim erisimi korunacak sekilde konumlandirildi.", "Musteri daha sessiz, daha stabil ve villa mimarisine daha uyumlu bir sogutma deneyimi elde etti.", "Benzer bir cozum icin kesif talep et", "iletisim.asp", "konyaalti-villa-klima-montaji", "case_konyaalti_villa_klima", 10)
    Call UygulamaCaseSeed("Mitsubishi Electric Klima Bakim Uygulamasi", "Bakim / Servis", "Performansi dusen bir Mitsubishi Electric klima icin filtre, drenaj ve calisma testlerini kapsayan periyodik bakim sureci.", "Muratpasa / Antalya", "Konut", "Klima bakimi", "Mitsubishi Electric split klima", "Ayni gun", "Klimada sogutma performansi dusmus, enerji tuketimi artmis ve ic ortam hava kalitesi zayiflamisti.", "Kesifte filtre kirliligi, drenaj akisi, serpantin durumu, gaz basinci ve calisma sesi kontrol edildi.", "Periyodik bakim planiyla filtre, serpantin ve drenaj hattinin temizlenmesi; gaz basinci ve calisma testlerinin tamamlanmasi onerildi.", "Ic ve dis unite temizlikleri yapildi, drenaj hatti acildi, gaz basinci kontrol edildi ve cihaz test modunda calistirildi.", "Daha dengeli sogutma, daha dusuk zorlanma ve daha konforlu ic ortam elde edildi.", "Bakim plani al", "iletisim.asp", "mitsubishi-electric-klima-bakim-uygulamasi", "case_mitsubishi_bakim", 20)
    Call UygulamaCaseSeed("Isyeri / Ofis VRF Sistem Cozumu", "VRF / Ticari", "Farkli odalarda farkli isi ihtiyaci olan ofis icin bolge bazli kontrol saglayan VRF sistem planlamasi.", "Muratpasa / Antalya", "Ofis", "VRF sistemleri", "Mitsubishi Electric VRF", "3 gun", "Farkli ofis bolumlerinde sicaklik dengesi saglanamiyor, toplanti odalari ve acik ofis alani ayni anda verimli yonetilemiyordu.", "Kesifte kullanim saatleri, bolumlerin kisi yogunlugu, cephe yonleri, tavan arasi gecisleri, otomasyon ihtiyaci ve dis unite konumu degerlendirildi.", "Bolge bazli ihtiyac analizi yapilarak uygun kapasitede VRF ic ve dis unite kombinasyonu planlandi.", "Boru hatti, otomasyon ayarlari ve oda bazli kontrol senaryolari devreye alindi. Teslimde kullaniciya temel kontrol egitimi verildi.", "Ofis genelinde dengeli iklimlendirme, kullanici bazli kontrol ve daha yonetilebilir enerji kullanimi elde edildi.", "VRF kesfi talep et", "iletisim.asp", "muratpasa-ofis-vrf-sistem-cozumu", "case_muratpasa_ofis_vrf", 30)
  End If
End Sub
%>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html>
<head>

<META HTTP-EQUIV="Content-Type" content="text/html; charset=windows-1254">
<meta name="viewport" content="width=device-width, initial-scale=1">

<title>Uygulama Ornekleri ve Case Study | Ozum Klima Antalya</title>
<meta name="description" content="Ozum Klima'nin Antalya'da tamamladigi klima montaji, VRF sistemleri, bakim ve havalandirma uygulamalarini detayli case study sayfalariyla inceleyin." />
<link rel="canonical" href="https://ozumklima.com/uygulama-ornekleri.asp" />

<meta name="author" content="Mega Tasarim"/>

<link rel="icon" type="image/png" href="fav.png" />
<link href="css/font.css" rel="stylesheet" type="text/css"/>
<link href="css/fontawesome-all.css" rel="stylesheet">
<script src="js/jquery-2.1.1.js"></script>
<script src="js/jquery-ui.min.js"></script>

<style>
  .case-page,
  .case-page * {
    box-sizing: border-box;
  }

  .case-page {
    width: 100%;
    background: #f7f7f7;
    padding: 0 0 58px 0;
    font-family: 'Source Sans Pro', Arial, sans-serif;
    color: #2c2c2c;
  }

  .case-hero {
    max-width: 1220px;
    margin: 0 auto;
    padding: 38px 18px 24px 18px;
  }

  .case-hero-card {
    position: relative;
    overflow: hidden;
    border: 1px solid #e6e0d3;
    border-radius: 22px;
    background: linear-gradient(135deg, #ffffff 0%, #f7f1df 100%);
    padding: 38px;
    box-shadow: 0 18px 45px rgba(34,34,34,.07);
  }

  .case-hero-card:after {
    content: "";
    position: absolute;
    right: -80px;
    top: -90px;
    width: 260px;
    height: 260px;
    border-radius: 50%;
    background: rgba(247,187,9,.22);
  }

  .case-kicker {
    display: inline-flex;
    align-items: center;
    gap: 8px;
    padding: 7px 12px;
    border-radius: 999px;
    background: #3c3c3c;
    color: #f7bb09;
    font-size: 13px;
    font-weight: 800;
    letter-spacing: .06em;
    text-transform: uppercase;
  }

  .case-title {
    position: relative;
    z-index: 1;
    margin: 18px 0 10px 0;
    max-width: 800px;
    font-size: 2.55em;
    line-height: 1.08em;
    color: #252525;
    font-weight: 800;
  }

  .case-subtitle {
    position: relative;
    z-index: 1;
    max-width: 760px;
    margin: 0;
    font-size: 1.12em;
    line-height: 1.65em;
    color: #5a5a5a;
  }

  .case-metrics {
    position: relative;
    z-index: 1;
    display: flex;
    flex-wrap: wrap;
    gap: 12px;
    margin-top: 24px;
  }

  .case-metric {
    min-width: 150px;
    padding: 14px 16px;
    border: 1px solid rgba(60,60,60,.10);
    border-radius: 15px;
    background: rgba(255,255,255,.78);
  }

  .case-metric strong {
    display: block;
    color: #262626;
    font-size: 1.25em;
  }

  .case-metric span {
    color: #6b6b6b;
    font-size: .92em;
  }

  .case-wrap {
    max-width: 1220px;
    margin: 0 auto;
    padding: 0 18px;
  }

  .case-grid {
    display: grid;
    grid-template-columns: repeat(3, minmax(0, 1fr));
    gap: 20px;
  }

  .case-card {
    display: flex;
    flex-direction: column;
    min-height: 100%;
    overflow: hidden;
    border: 1px solid #e4e4e4;
    border-radius: 20px;
    background: #ffffff;
    box-shadow: 0 14px 34px rgba(0,0,0,.055);
    transition: transform .18s ease, box-shadow .18s ease, border-color .18s ease;
  }

  .case-card:hover {
    transform: translateY(-3px);
    border-color: #f7bb09;
    box-shadow: 0 18px 42px rgba(0,0,0,.085);
  }

  .case-cover {
    position: relative;
    height: 210px;
    background: linear-gradient(135deg, #353535, #595959);
    overflow: hidden;
  }

  .case-cover img {
    width: 100%;
    height: 100%;
    object-fit: cover;
    display: block;
  }

  .case-cover-placeholder {
    display: flex;
    align-items: center;
    justify-content: center;
    height: 100%;
    padding: 22px;
    color: #f7bb09;
    font-size: 1.1em;
    font-weight: 800;
    text-align: center;
  }

  .case-card-body {
    display: flex;
    flex-direction: column;
    flex: 1;
    padding: 22px;
  }

  .case-badge {
    display: inline-flex;
    align-self: flex-start;
    margin-bottom: 14px;
    padding: 6px 10px;
    border-radius: 999px;
    background: #fff7dd;
    color: #7c5a00;
    font-weight: 800;
    font-size: 12px;
    letter-spacing: .03em;
    text-transform: uppercase;
  }

  .case-card h2 {
    margin: 0 0 10px 0;
    font-size: 1.36em;
    line-height: 1.28em;
    color: #232323;
    font-weight: 800;
  }

  .case-meta {
    display: flex;
    flex-wrap: wrap;
    gap: 7px;
    margin: 0 0 14px 0;
  }

  .case-meta span {
    border-radius: 999px;
    background: #f3f3f3;
    color: #555;
    padding: 5px 9px;
    font-size: 12px;
    font-weight: 700;
  }

  .case-excerpt {
    margin: 0 0 18px 0;
    color: #525252;
    font-size: 1em;
    line-height: 1.58em;
  }

  .case-card-link {
    margin-top: auto;
  }

  .case-card-link a {
    display: inline-flex;
    align-items: center;
    justify-content: center;
    gap: 8px;
    width: 100%;
    min-height: 44px;
    padding: 11px 15px;
    border: 1px solid #2f2f2f;
    border-radius: 999px;
    background: #2f2f2f;
    color: #ffffff;
    text-decoration: none;
    font-weight: 800;
    transition: background .18s ease, color .18s ease, border-color .18s ease;
  }

  .case-card-link a:hover {
    background: #f7bb09;
    border-color: #f7bb09;
    color: #232323;
  }

  .case-empty {
    padding: 28px;
    border: 1px dashed #d9d9d9;
    border-radius: 16px;
    background: #ffffff;
    color: #555;
    font-size: 1.05em;
  }

  @media screen and (max-width: 980px) {
    .case-grid {
      grid-template-columns: repeat(2, minmax(0, 1fr));
    }

    .case-title {
      font-size: 2.1em;
    }
  }

  @media screen and (max-width: 640px) {
    .case-hero-card {
      padding: 24px 20px;
      border-radius: 16px;
    }

    .case-grid {
      grid-template-columns: 1fr;
    }

    .case-title {
      font-size: 1.75em;
    }
  }
</style>

</head>

<body>

<!--#include file='ust.asp'-->

<%
Call UygulamaCaseEnsureTable()

Dim uygulamaRs, uygulamaSql, uygulamaSayisi
Dim cardSlug, cardDetailUrl, cardCover, cardCoverAlt, cardExcerpt
Dim cardLokasyon, cardProjeTipi, cardHizmetTipi, cardSistemMarka
uygulamaSayisi = 0
Set uygulamaRs = Server.CreateObject("ADODB.RecordSet")
uygulamaSql = "SELECT * FROM uygulama_ornekleri WHERE [yayin]<>0 ORDER BY [sira] ASC, [id] ASC"
uygulamaRs.Open uygulamaSql, baglanti, 1, 1
%>

<div class="case-page">
  <section class="case-hero">
    <div class="case-hero-card">
      <div class="case-kicker">Saha notlari</div>
      <h1 class="case-title">Uygulama Ornekleri</h1>
      <p class="case-subtitle">Ozum Klima'nin Antalya'da hayata gecirdigi gercek saha uygulamalarini kart ozetleriyle inceleyin. Her ornek kendi detay sayfasinda problem, kesif, cozum, uygulama, sonuc, fotograf ve CTA akisiyle anlatilir.</p>

      <div class="case-metrics">
        <div class="case-metric"><strong>Kesif</strong><span>Alan, kapasite ve ihtiyac analizi</span></div>
        <div class="case-metric"><strong>Kanit</strong><span>Foto, galeri ve opsiyonel video</span></div>
        <div class="case-metric"><strong>Sonuc</strong><span>Okunabilir case study sayfasi</span></div>
      </div>
    </div>
  </section>

  <div class="case-wrap">
    <% If uygulamaRs.EOF Then %>
      <div class="case-empty">Yayinda uygulama ornegi bulunamadi. Panelden yeni bir kayit ekleyebilirsiniz.</div>
    <% Else %>
      <div class="case-grid">
        <%
        Do While Not uygulamaRs.EOF
          uygulamaSayisi = uygulamaSayisi + 1
          cardSlug = UygulamaCaseValue(uygulamaRs, "slug")
          cardDetailUrl = UygulamaCaseDetailUrl(cardSlug)
          cardCover = UygulamaCaseValue(uygulamaRs, "cover_image")
          cardCoverAlt = UygulamaCaseValue(uygulamaRs, "cover_alt")
          If Trim(cardCoverAlt) = "" Then cardCoverAlt = UygulamaCaseValue(uygulamaRs, "baslik")
          cardExcerpt = UygulamaCaseValue(uygulamaRs, "kisa_ozet")
          If Trim(cardExcerpt) = "" Then cardExcerpt = UygulamaCaseValue(uygulamaRs, "problem")
          cardLokasyon = UygulamaCaseValue(uygulamaRs, "lokasyon")
          cardProjeTipi = UygulamaCaseValue(uygulamaRs, "proje_tipi")
          cardHizmetTipi = UygulamaCaseValue(uygulamaRs, "hizmet_tipi")
          cardSistemMarka = UygulamaCaseValue(uygulamaRs, "sistem_marka")
        %>
          <article class="case-card">
            <a class="case-cover" href="<%=cardDetailUrl%>" aria-label="<%=UygulamaCaseAttr(uygulamaRs("baslik"))%> detay sayfasini ac">
              <% If Trim(cardCover) <> "" Then %>
                <img src="<%=UygulamaCaseImage(cardCover)%>" alt="<%=UygulamaCaseAttr(cardCoverAlt)%>">
              <% Else %>
                <div class="case-cover-placeholder">Ozum Klima<br>Uygulama Ornegi</div>
              <% End If %>
            </a>

            <div class="case-card-body">
              <% If Trim(uygulamaRs("etiket") & "") <> "" Then %>
                <div class="case-badge"><%=UygulamaCaseText(uygulamaRs("etiket"))%></div>
              <% End If %>

              <h2><%=UygulamaCaseText(uygulamaRs("baslik"))%></h2>

              <div class="case-meta">
                <% If Trim(cardLokasyon) <> "" Then %><span><%=UygulamaCaseText(cardLokasyon)%></span><% End If %>
                <% If Trim(cardProjeTipi) <> "" Then %><span><%=UygulamaCaseText(cardProjeTipi)%></span><% End If %>
                <% If Trim(cardHizmetTipi) <> "" Then %><span><%=UygulamaCaseText(cardHizmetTipi)%></span><% End If %>
                <% If Trim(cardSistemMarka) <> "" Then %><span><%=UygulamaCaseText(cardSistemMarka)%></span><% End If %>
              </div>

              <% If Trim(cardExcerpt) <> "" Then %>
                <p class="case-excerpt"><%=UygulamaCaseText(cardExcerpt)%></p>
              <% End If %>

              <div class="case-card-link">
                <a href="<%=cardDetailUrl%>" onclick="if(window.analyticsTrack){analyticsTrack('case_study_detail_click', {page: 'uygulama-ornekleri', case_slug: '<%=UygulamaCaseJs(cardSlug)%>'});}">
                  Case'i oku <span>&#8594;</span>
                </a>
              </div>
            </div>
          </article>
        <%
          uygulamaRs.MoveNext
        Loop
        %>
      </div>
    <% End If %>
  </div>
</div>

<%
If uygulamaRs.State = 1 Then uygulamaRs.Close
Set uygulamaRs = Nothing
%>

<!--  #include file='alt.asp' -->

</body>
</html>
