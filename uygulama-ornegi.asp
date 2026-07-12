<!--#include file='db.asp'-->
<%
Function UygulamaDetailSql(ByVal value)
  UygulamaDetailSql = Replace(CStr(value & ""), "'", "''")
End Function

Function UygulamaDetailText(ByVal value)
  Dim textValue
  textValue = Server.HTMLEncode(CStr(value & ""))
  textValue = Replace(textValue, vbCrLf, "<br>")
  textValue = Replace(textValue, vbLf, "<br>")
  UygulamaDetailText = textValue
End Function

Function UygulamaDetailAttr(ByVal value)
  UygulamaDetailAttr = Server.HTMLEncode(CStr(value & ""))
End Function

Function UygulamaDetailParagraphs(ByVal value)
  Dim textValue, htmlValue, parts, i, onePart
  textValue = Trim(CStr(value & ""))
  If textValue = "" Then
    UygulamaDetailParagraphs = ""
    Exit Function
  End If

  textValue = Replace(textValue, vbCrLf, vbLf)
  textValue = Replace(textValue, vbCr, vbLf)
  parts = Split(textValue, vbLf)
  htmlValue = ""

  For i = 0 To UBound(parts)
    onePart = Trim(parts(i))
    If onePart <> "" Then
      htmlValue = htmlValue & "<p>" & Server.HTMLEncode(onePart) & "</p>"
    End If
  Next

  If htmlValue = "" Then htmlValue = "<p>" & Server.HTMLEncode(textValue) & "</p>"
  UygulamaDetailParagraphs = htmlValue
End Function

Function UygulamaDetailValue(ByRef rs, ByVal fieldName)
  On Error Resume Next
  UygulamaDetailValue = rs(fieldName) & ""
  If Err.Number <> 0 Then
    UygulamaDetailValue = ""
    Err.Clear
  End If
  On Error GoTo 0
End Function

Function UygulamaDetailUrl(ByVal value)
  Dim urlValue, lowerUrl
  urlValue = Trim(CStr(value & ""))
  If urlValue = "" Then urlValue = "iletisim.asp"
  If Left(urlValue, 3) = "../" Then urlValue = Mid(urlValue, 4)

  lowerUrl = LCase(urlValue)
  If InStr(1, lowerUrl, "javascript:", 1) > 0 Then urlValue = "iletisim.asp"

  UygulamaDetailUrl = Server.HTMLEncode(urlValue)
End Function

Function UygulamaDetailImage(ByVal value)
  Dim imageValue, lowerImage
  imageValue = Trim(CStr(value & ""))
  lowerImage = LCase(imageValue)
  If InStr(1, lowerImage, "javascript:", 1) > 0 Then imageValue = ""
  UygulamaDetailImage = Server.HTMLEncode(imageValue)
End Function

Function UygulamaDetailVideoEmbed(ByVal value)
  Dim videoValue, lowerVideo, pos, videoId, ampPos
  videoValue = Trim(CStr(value & ""))
  lowerVideo = LCase(videoValue)

  If videoValue = "" Then
    UygulamaDetailVideoEmbed = ""
    Exit Function
  End If

  If InStr(1, lowerVideo, "javascript:", 1) > 0 Then
    UygulamaDetailVideoEmbed = ""
    Exit Function
  End If

  If InStr(1, lowerVideo, "youtube.com/watch", 1) > 0 And InStr(1, videoValue, "v=", 1) > 0 Then
    pos = InStr(1, videoValue, "v=", 1) + 2
    videoId = Mid(videoValue, pos)
    ampPos = InStr(1, videoId, "&", 1)
    If ampPos > 0 Then videoId = Left(videoId, ampPos - 1)
    UygulamaDetailVideoEmbed = "https://www.youtube.com/embed/" & Server.HTMLEncode(videoId)
    Exit Function
  End If

  If InStr(1, lowerVideo, "youtu.be/", 1) > 0 Then
    pos = InStr(1, lowerVideo, "youtu.be/", 1) + 9
    videoId = Mid(videoValue, pos)
    ampPos = InStr(1, videoId, "?", 1)
    If ampPos > 0 Then videoId = Left(videoId, ampPos - 1)
    UygulamaDetailVideoEmbed = "https://www.youtube.com/embed/" & Server.HTMLEncode(videoId)
    Exit Function
  End If

  If InStr(1, lowerVideo, "vimeo.com/", 1) > 0 Then
    pos = InStrRev(videoValue, "/") + 1
    videoId = Mid(videoValue, pos)
    ampPos = InStr(1, videoId, "?", 1)
    If ampPos > 0 Then videoId = Left(videoId, ampPos - 1)
    UygulamaDetailVideoEmbed = "https://player.vimeo.com/video/" & Server.HTMLEncode(videoId)
    Exit Function
  End If

  UygulamaDetailVideoEmbed = ""
End Function

Sub UygulamaDetailAddColumn(ByVal columnSql)
  On Error Resume Next
  baglanti.Execute "ALTER TABLE uygulama_ornekleri ADD COLUMN " & columnSql
  Err.Clear
  On Error GoTo 0
End Sub

Sub UygulamaDetailSeed(ByVal baslik, ByVal etiket, ByVal kisaOzet, ByVal lokasyon, ByVal projeTipi, ByVal hizmetTipi, ByVal sistemMarka, ByVal projeSuresi, ByVal problem, ByVal kesif, ByVal cozum, ByVal uygulama, ByVal sonuc, ByVal ctaMetni, ByVal ctaUrl, ByVal slug, ByVal takipEtiketi, ByVal sira)
  Dim seedSql
  seedSql = "INSERT INTO uygulama_ornekleri ([baslik], [etiket], [kisa_ozet], [lokasyon], [proje_tipi], [hizmet_tipi], [sistem_marka], [proje_suresi], [problem], [kesif_degerlendirme], [cozum], [uygulama], [sonuc], [cta_metni], [cta_url], [slug], [takip_etiketi], [sira], [yayin], [kayit_tarihi]) VALUES (" & _
            "'" & UygulamaDetailSql(baslik) & "'," & _
            "'" & UygulamaDetailSql(etiket) & "'," & _
            "'" & UygulamaDetailSql(kisaOzet) & "'," & _
            "'" & UygulamaDetailSql(lokasyon) & "'," & _
            "'" & UygulamaDetailSql(projeTipi) & "'," & _
            "'" & UygulamaDetailSql(hizmetTipi) & "'," & _
            "'" & UygulamaDetailSql(sistemMarka) & "'," & _
            "'" & UygulamaDetailSql(projeSuresi) & "'," & _
            "'" & UygulamaDetailSql(problem) & "'," & _
            "'" & UygulamaDetailSql(kesif) & "'," & _
            "'" & UygulamaDetailSql(cozum) & "'," & _
            "'" & UygulamaDetailSql(uygulama) & "'," & _
            "'" & UygulamaDetailSql(sonuc) & "'," & _
            "'" & UygulamaDetailSql(ctaMetni) & "'," & _
            "'" & UygulamaDetailSql(ctaUrl) & "'," & _
            "'" & UygulamaDetailSql(slug) & "'," & _
            "'" & UygulamaDetailSql(takipEtiketi) & "'," & CLng(sira) & ",-1,Now())"
  baglanti.Execute seedSql
End Sub

Sub UygulamaDetailEnsureTable()
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

  Call UygulamaDetailAddColumn("[kisa_ozet] MEMO")
  Call UygulamaDetailAddColumn("[lokasyon] TEXT(150)")
  Call UygulamaDetailAddColumn("[proje_tipi] TEXT(120)")
  Call UygulamaDetailAddColumn("[hizmet_tipi] TEXT(150)")
  Call UygulamaDetailAddColumn("[sistem_marka] TEXT(180)")
  Call UygulamaDetailAddColumn("[proje_suresi] TEXT(80)")
  Call UygulamaDetailAddColumn("[kesif_degerlendirme] MEMO")
  Call UygulamaDetailAddColumn("[cover_image] TEXT(255)")
  Call UygulamaDetailAddColumn("[cover_alt] TEXT(255)")
  Call UygulamaDetailAddColumn("[galeri_gorseller] MEMO")
  Call UygulamaDetailAddColumn("[galeri_altlar] MEMO")
  Call UygulamaDetailAddColumn("[video_url] TEXT(255)")
  Call UygulamaDetailAddColumn("[oncesi_image] TEXT(255)")
  Call UygulamaDetailAddColumn("[oncesi_alt] TEXT(255)")
  Call UygulamaDetailAddColumn("[sonrasi_image] TEXT(255)")
  Call UygulamaDetailAddColumn("[sonrasi_alt] TEXT(255)")
  Call UygulamaDetailAddColumn("[musteri_yorumu] MEMO")
  Call UygulamaDetailAddColumn("[sss] MEMO")
  Call UygulamaDetailAddColumn("[ilgili_hizmet_url] TEXT(255)")
  Call UygulamaDetailAddColumn("[ilgili_bolge_url] TEXT(255)")
  Call UygulamaDetailAddColumn("[takip_etiketi] TEXT(120)")
  Call UygulamaDetailAddColumn("[seo_title] TEXT(255)")
  Call UygulamaDetailAddColumn("[meta_description] MEMO")


  
  On Error Resume Next
  Set testRs = baglanti.Execute("SELECT COUNT(*) AS toplam FROM uygulama_ornekleri")
  If Err.Number = 0 Then toplamKayit = CLng(testRs("toplam"))
  If Not testRs Is Nothing Then
    If testRs.State = 1 Then testRs.Close
    Set testRs = Nothing
  End If
  On Error GoTo 0

  If toplamKayit = 0 Then
    Call UygulamaDetailSeed("Antalya Villa Klima Montaji", "Villa / Konut", "Konyaalti bolgesindeki villada salon ve yatak odalari icin sessiz, dengeli ve estetik klima cozumunun nasil planlandigini inceleyin.", "Konyaalti / Antalya", "Villa", "Klima montaji", "Mitsubishi Electric multi split", "1 gun", "Musteri yaz aylarinda salon ve yatak odalarinda yeterli sogutma alamiyor, dis unite konumunun goruntu ve ses acisindan sorun yaratmamasini istiyordu.", "Kesifte oda metrekareleri, cephe yonu, cam alani, drenaj hattinin gececegi guzergah, elektrik altyapisi ve dis unite servis erisimi kontrol edildi.", "Alan ihtiyacina uygun kapasite hesaplandi ve birden fazla ic uniteyi dengeli calistirabilecek Mitsubishi Electric cozum onerildi.", "Ic unite yerleri musteriyle netlestirildi, boru ve drenaj hatlari gizli guzergahlarla tasindi, dis unite bakim erisimi korunacak sekilde konumlandirildi.", "Musteri daha sessiz, daha stabil ve villa mimarisine daha uyumlu bir sogutma deneyimi elde etti.", "Benzer bir cozum icin kesif talep et", "iletisim.asp", "konyaalti-villa-klima-montaji", "case_konyaalti_villa_klima", 10)
  End If
End Sub

Sub UygulamaRenderInfo(ByVal labelText, ByVal valueText)
  If Trim(CStr(valueText & "")) <> "" Then
    Response.Write "<div class=""detail-info""><span>" & Server.HTMLEncode(labelText) & "</span><strong>" & UygulamaDetailText(valueText) & "</strong></div>"
  End If
End Sub
%>

<%
Call UygulamaDetailEnsureTable()

Dim slug, detailRs, detailSql, pageFound
Dim dBaslik, dEtiket, dKisaOzet, dLokasyon, dProjeTipi, dHizmetTipi, dSistemMarka, dSure
Dim dProblem, dKesif, dCozum, dUygulama, dSonuc, dCover, dCoverAlt, dGallery, dGalleryAlts
Dim dVideo, dBefore, dBeforeAlt, dAfter, dAfterAlt, dYorum, dFaq, dCtaMetni, dCtaUrl, dTakip
Dim dSeoTitle, dMetaDesc, dHizmetUrl, dBolgeUrl

slug = Trim(Request.QueryString("slug") & "")
If slug = "" Then Response.Redirect "uygulama-ornekleri.asp"

Set detailRs = Server.CreateObject("ADODB.RecordSet")
detailSql = "SELECT TOP 1 * FROM uygulama_ornekleri WHERE [yayin]<>0 AND [slug]='" & UygulamaDetailSql(slug) & "'"
detailRs.Open detailSql, baglanti, 1, 1

pageFound = False
If Not detailRs.EOF Then
  pageFound = True
  dBaslik = detailRs("baslik") & ""
  dEtiket = detailRs("etiket") & ""
  dKisaOzet = UygulamaDetailValue(detailRs, "kisa_ozet")
  dLokasyon = UygulamaDetailValue(detailRs, "lokasyon")
  dProjeTipi = UygulamaDetailValue(detailRs, "proje_tipi")
  dHizmetTipi = UygulamaDetailValue(detailRs, "hizmet_tipi")
  dSistemMarka = UygulamaDetailValue(detailRs, "sistem_marka")
  dSure = UygulamaDetailValue(detailRs, "proje_suresi")
  dProblem = UygulamaDetailValue(detailRs, "problem")
  dKesif = UygulamaDetailValue(detailRs, "kesif_degerlendirme")
  dCozum = UygulamaDetailValue(detailRs, "cozum")
  dUygulama = UygulamaDetailValue(detailRs, "uygulama")
  dSonuc = UygulamaDetailValue(detailRs, "sonuc")
  dCover = UygulamaDetailValue(detailRs, "cover_image")
  dCoverAlt = UygulamaDetailValue(detailRs, "cover_alt")
  dGallery = UygulamaDetailValue(detailRs, "galeri_gorseller")
  dGalleryAlts = UygulamaDetailValue(detailRs, "galeri_altlar")
  dVideo = UygulamaDetailValue(detailRs, "video_url")
  dBefore = UygulamaDetailValue(detailRs, "oncesi_image")
  dBeforeAlt = UygulamaDetailValue(detailRs, "oncesi_alt")
  dAfter = UygulamaDetailValue(detailRs, "sonrasi_image")
  dAfterAlt = UygulamaDetailValue(detailRs, "sonrasi_alt")
  dYorum = UygulamaDetailValue(detailRs, "musteri_yorumu")
  dFaq = UygulamaDetailValue(detailRs, "sss")
  dCtaMetni = UygulamaDetailValue(detailRs, "cta_metni")
  dCtaUrl = UygulamaDetailValue(detailRs, "cta_url")
  dTakip = UygulamaDetailValue(detailRs, "takip_etiketi")
  dSeoTitle = UygulamaDetailValue(detailRs, "seo_title")
  dMetaDesc = UygulamaDetailValue(detailRs, "meta_description")
  dHizmetUrl = UygulamaDetailValue(detailRs, "ilgili_hizmet_url")
  dBolgeUrl = UygulamaDetailValue(detailRs, "ilgili_bolge_url")
End If

If detailRs.State = 1 Then detailRs.Close
Set detailRs = Nothing

If Not pageFound Then Response.Redirect "uygulama-ornekleri.asp"
If Trim(dSeoTitle) = "" Then dSeoTitle = dBaslik & " | Ozum Klima Antalya"
If Trim(dMetaDesc) = "" Then dMetaDesc = dKisaOzet
If Trim(dMetaDesc) = "" Then dMetaDesc = dProblem
If Trim(dCtaMetni) = "" Then dCtaMetni = "Benzer bir proje icin kesif talep et"
If Trim(dCtaUrl) = "" Then dCtaUrl = "iletisim.asp"
If Trim(dTakip) = "" Then dTakip = slug
If Trim(dCoverAlt) = "" Then dCoverAlt = dBaslik
If Trim(dBeforeAlt) = "" Then dBeforeAlt = dBaslik & " oncesi"
If Trim(dAfterAlt) = "" Then dAfterAlt = dBaslik & " sonrasi"
%>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html>
<head>

<META HTTP-EQUIV="Content-Type" content="text/html; charset=windows-1254">
<meta name="viewport" content="width=device-width, initial-scale=1">

<title><%=UygulamaDetailAttr(dSeoTitle)%></title>
<meta name="description" content="<%=UygulamaDetailAttr(dMetaDesc)%>" />
<link rel="canonical" href="https://ozumklima.com/uygulama-ornegi.asp?slug=<%=Server.URLEncode(slug)%>" />

<meta name="author" content="Mega Tasarim"/>

<link rel="icon" type="image/png" href="fav.png" />
<link href="css/font.css" rel="stylesheet" type="text/css"/>
<link href="css/fontawesome-all.css" rel="stylesheet">
<script src="js/jquery-2.1.1.js"></script>
<script src="js/jquery-ui.min.js"></script>

<style>
  .detail-page,
  .detail-page * {
    box-sizing: border-box;
  }

  .detail-page {
    background: #f6f6f4;
    color: #242424;
    font-family: 'Source Sans Pro', Arial, sans-serif;
    padding-bottom: 58px;
  }

  .detail-shell {
    max-width: 1120px;
    margin: 0 auto;
    padding: 0 18px;
  }

  .detail-back {
    padding-top: 28px;
  }

  .detail-back a {
    display: inline-flex;
    gap: 8px;
    align-items: center;
    color: #4b4b4b;
    text-decoration: none;
    font-weight: 800;
  }

  .detail-hero {
    padding: 24px 0 20px 0;
  }

  .detail-kicker {
    display: inline-flex;
    align-items: center;
    padding: 7px 12px;
    border-radius: 999px;
    background: #2f2f2f;
    color: #f7bb09;
    font-size: 13px;
    font-weight: 800;
    letter-spacing: .06em;
    text-transform: uppercase;
  }

  .detail-title {
    max-width: 900px;
    margin: 18px 0 14px 0;
    color: #202020;
    font-size: 3em;
    line-height: 1.08em;
    font-weight: 900;
  }

  .detail-lead {
    max-width: 780px;
    margin: 0;
    color: #555;
    font-size: 1.24em;
    line-height: 1.72em;
  }

  .detail-info-grid {
    display: grid;
    grid-template-columns: repeat(4, minmax(0, 1fr));
    gap: 12px;
    margin-top: 28px;
  }

  .detail-info {
    border: 1px solid #e1dfd8;
    border-radius: 16px;
    background: #ffffff;
    padding: 15px;
  }

  .detail-info span {
    display: block;
    color: #777;
    font-size: 12px;
    font-weight: 800;
    letter-spacing: .06em;
    text-transform: uppercase;
    margin-bottom: 5px;
  }

  .detail-info strong {
    display: block;
    color: #262626;
    font-size: 1.02em;
    line-height: 1.35em;
  }

  .detail-cover {
    overflow: hidden;
    margin: 10px 0 34px 0;
    border-radius: 24px;
    background: linear-gradient(135deg, #2f2f2f, #575757);
    min-height: 320px;
    box-shadow: 0 18px 42px rgba(0,0,0,.08);
  }

  .detail-cover img {
    display: block;
    width: 100%;
    max-height: 560px;
    object-fit: cover;
  }

  .detail-cover-placeholder {
    min-height: 340px;
    display: flex;
    align-items: center;
    justify-content: center;
    padding: 34px;
    color: #f7bb09;
    font-size: 1.4em;
    font-weight: 900;
    text-align: center;
  }

  .detail-layout {
    display: grid;
    grid-template-columns: minmax(0, 1fr) 310px;
    gap: 28px;
    align-items: start;
  }

  .detail-article {
    border: 1px solid #e4e0d8;
    border-radius: 24px;
    background: #ffffff;
    padding: 34px;
    box-shadow: 0 14px 34px rgba(0,0,0,.045);
  }

  .detail-section {
    padding: 0 0 30px 0;
    margin-bottom: 28px;
    border-bottom: 1px solid #eee9df;
  }

  .detail-section:last-child {
    margin-bottom: 0;
    padding-bottom: 0;
    border-bottom: 0;
  }

  .detail-section h2 {
    margin: 0 0 13px 0;
    color: #202020;
    font-size: 1.75em;
    line-height: 1.24em;
  }

  .detail-section p {
    margin: 0 0 14px 0;
    color: #3f3f3f;
    font-size: 1.12em;
    line-height: 1.82em;
  }

  .detail-section p:last-child {
    margin-bottom: 0;
  }

  .detail-quote {
    border-left: 5px solid #f7bb09;
    border-radius: 14px;
    background: #fff8e5;
    padding: 18px 20px;
    color: #4a3a0a;
    font-size: 1.08em;
    line-height: 1.7em;
    font-weight: 700;
  }

  .detail-gallery {
    display: grid;
    grid-template-columns: repeat(2, minmax(0, 1fr));
    gap: 14px;
  }

  .detail-gallery figure,
  .detail-before-after figure {
    margin: 0;
    border: 1px solid #ececec;
    border-radius: 18px;
    overflow: hidden;
    background: #fafafa;
  }

  .detail-gallery img,
  .detail-before-after img {
    display: block;
    width: 100%;
    height: 255px;
    object-fit: cover;
  }

  .detail-gallery figcaption,
  .detail-before-after figcaption {
    padding: 11px 13px;
    color: #5c5c5c;
    font-size: 13px;
    font-weight: 700;
  }

  .detail-before-after {
    display: grid;
    grid-template-columns: repeat(2, minmax(0, 1fr));
    gap: 14px;
  }

  .detail-video-frame {
    position: relative;
    overflow: hidden;
    border-radius: 18px;
    background: #222;
    padding-top: 56.25%;
  }

  .detail-video-frame iframe {
    position: absolute;
    inset: 0;
    width: 100%;
    height: 100%;
    border: 0;
  }

  .detail-video-link {
    display: inline-flex;
    align-items: center;
    justify-content: center;
    margin-top: 10px;
    padding: 11px 16px;
    border-radius: 999px;
    background: #2f2f2f;
    color: #ffffff;
    text-decoration: none;
    font-weight: 800;
  }

  .detail-faq-item {
    border: 1px solid #ece5d6;
    border-radius: 16px;
    padding: 16px;
    margin-bottom: 12px;
    background: #fffdf7;
  }

  .detail-faq-item h3 {
    margin: 0 0 8px 0;
    font-size: 1.12em;
    color: #262626;
  }

  .detail-faq-item p {
    font-size: 1em;
    line-height: 1.65em;
  }

  .detail-side {
    position: sticky;
    top: 18px;
  }

  .detail-cta {
    border: 1px solid #e5dcc8;
    border-radius: 22px;
    background: linear-gradient(135deg, #ffffff, #fff4d6);
    padding: 22px;
    box-shadow: 0 14px 30px rgba(0,0,0,.055);
  }

  .detail-cta h2 {
    margin: 0 0 10px 0;
    color: #222;
    font-size: 1.35em;
    line-height: 1.25em;
  }

  .detail-cta p {
    margin: 0 0 16px 0;
    color: #555;
    line-height: 1.6em;
  }

  .detail-cta a {
    display: inline-flex;
    justify-content: center;
    align-items: center;
    gap: 8px;
    width: 100%;
    min-height: 46px;
    border-radius: 999px;
    background: #2f2f2f;
    color: #ffffff;
    text-decoration: none;
    font-weight: 900;
  }

  .detail-side-links {
    margin-top: 14px;
    border: 1px solid #e8e8e8;
    border-radius: 18px;
    background: #ffffff;
    padding: 16px;
  }

  .detail-side-links a {
    display: block;
    margin-top: 9px;
    color: #333;
    font-weight: 800;
    text-decoration: none;
  }

  @media screen and (max-width: 980px) {
    .detail-title {
      font-size: 2.35em;
    }

    .detail-info-grid {
      grid-template-columns: repeat(2, minmax(0, 1fr));
    }

    .detail-layout {
      grid-template-columns: 1fr;
    }

    .detail-side {
      position: static;
    }
  }

  @media screen and (max-width: 640px) {
    .detail-title {
      font-size: 1.95em;
    }

    .detail-info-grid,
    .detail-gallery,
    .detail-before-after {
      grid-template-columns: 1fr;
    }

    .detail-article {
      padding: 24px 20px;
      border-radius: 18px;
    }
  }
</style>

</head>

<body>

<!--#include file='ust.asp'-->

<div class="detail-page">
  <div class="detail-shell">
    <div class="detail-back"><a href="uygulama-ornekleri.asp">&#8592; Uygulama orneklerine don</a></div>

    <section class="detail-hero">
      <% If Trim(dEtiket) <> "" Then %><div class="detail-kicker"><%=UygulamaDetailText(dEtiket)%></div><% End If %>
      <h1 class="detail-title"><%=UygulamaDetailText(dBaslik)%></h1>
      <% If Trim(dKisaOzet) <> "" Then %><p class="detail-lead"><%=UygulamaDetailText(dKisaOzet)%></p><% End If %>

      <div class="detail-info-grid">
        <% Call UygulamaRenderInfo("Lokasyon", dLokasyon) %>
        <% Call UygulamaRenderInfo("Proje tipi", dProjeTipi) %>
        <% Call UygulamaRenderInfo("Hizmet", dHizmetTipi) %>
        <% Call UygulamaRenderInfo("Sistem / marka", dSistemMarka) %>
        <% Call UygulamaRenderInfo("Sure", dSure) %>
      </div>
    </section>

    <div class="detail-cover">
      <% If Trim(dCover) <> "" Then %>
        <img src="<%=UygulamaDetailImage(dCover)%>" alt="<%=UygulamaDetailAttr(dCoverAlt)%>">
      <% Else %>
        <div class="detail-cover-placeholder">Ozum Klima uygulama ornegi</div>
      <% End If %>
    </div>

    <div class="detail-layout">
      <article class="detail-article">
        <% If Trim(dProblem) <> "" Then %>
          <section class="detail-section">
            <h2>Problem</h2>
            <%=UygulamaDetailParagraphs(dProblem)%>
          </section>
        <% End If %>

        <% If Trim(dKesif) <> "" Then %>
          <section class="detail-section">
            <h2>Kesif ve teknik degerlendirme</h2>
            <%=UygulamaDetailParagraphs(dKesif)%>
          </section>
        <% End If %>

        <% If Trim(dCozum) <> "" Then %>
          <section class="detail-section">
            <h2>Cozum</h2>
            <%=UygulamaDetailParagraphs(dCozum)%>
          </section>
        <% End If %>

        <% If Trim(dUygulama) <> "" Then %>
          <section class="detail-section">
            <h2>Uygulama sureci</h2>
            <%=UygulamaDetailParagraphs(dUygulama)%>
          </section>
        <% End If %>

        <% If Trim(dSonuc) <> "" Then %>
          <section class="detail-section">
            <h2>Sonuc</h2>
            <%=UygulamaDetailParagraphs(dSonuc)%>
          </section>
        <% End If %>

        <% If Trim(dYorum) <> "" Then %>
          <section class="detail-section">
            <h2>Musteri notu</h2>
            <div class="detail-quote"><%=UygulamaDetailText(dYorum)%></div>
          </section>
        <% End If %>

        <% If Trim(dBefore) <> "" Or Trim(dAfter) <> "" Then %>
          <section class="detail-section">
            <h2>Oncesi / sonrasi</h2>
            <div class="detail-before-after">
              <% If Trim(dBefore) <> "" Then %>
                <figure>
                  <img src="<%=UygulamaDetailImage(dBefore)%>" alt="<%=UygulamaDetailAttr(dBeforeAlt)%>">
                  <figcaption>Oncesi</figcaption>
                </figure>
              <% End If %>
              <% If Trim(dAfter) <> "" Then %>
                <figure>
                  <img src="<%=UygulamaDetailImage(dAfter)%>" alt="<%=UygulamaDetailAttr(dAfterAlt)%>">
                  <figcaption>Sonrasi</figcaption>
                </figure>
              <% End If %>
            </div>
          </section>
        <% End If %>

        <% If Trim(dGallery) <> "" Then %>
          <section class="detail-section">
            <h2>Foto galeri</h2>
            <div class="detail-gallery">
              <%
              Dim galleryLines, galleryAltLines, g, gUrl, gAlt
              galleryLines = Split(Replace(Replace(dGallery, vbCrLf, vbLf), vbCr, vbLf), vbLf)
              galleryAltLines = Split(Replace(Replace(dGalleryAlts, vbCrLf, vbLf), vbCr, vbLf), vbLf)
              For g = 0 To UBound(galleryLines)
                gUrl = Trim(galleryLines(g))
                If g <= UBound(galleryAltLines) Then
                  gAlt = Trim(galleryAltLines(g))
                Else
                  gAlt = ""
                End If
                If gAlt = "" Then gAlt = dBaslik
                If gUrl <> "" Then
              %>
                <figure>
                  <img src="<%=UygulamaDetailImage(gUrl)%>" alt="<%=UygulamaDetailAttr(gAlt)%>">
                  <figcaption><%=UygulamaDetailText(gAlt)%></figcaption>
                </figure>
              <%
                End If
              Next
              %>
            </div>
          </section>
        <% End If %>

        <% If Trim(dVideo) <> "" Then %>
          <section class="detail-section">
            <h2>Video</h2>
            <% If Trim(UygulamaDetailVideoEmbed(dVideo)) <> "" Then %>
              <div class="detail-video-frame">
                <iframe src="<%=UygulamaDetailVideoEmbed(dVideo)%>" title="<%=UygulamaDetailAttr(dBaslik)%> video" allowfullscreen></iframe>
              </div>
            <% Else %>
              <a class="detail-video-link" href="<%=UygulamaDetailUrl(dVideo)%>" target="_blank" rel="noopener">Videoyu ac</a>
            <% End If %>
          </section>
        <% End If %>

        <% If Trim(dFaq) <> "" Then %>
          <section class="detail-section">
            <h2>Sik sorulan sorular</h2>
            <%
            Dim faqLines, f, faqLine, pipePos, faqQuestion, faqAnswer
            faqLines = Split(Replace(Replace(dFaq, vbCrLf, vbLf), vbCr, vbLf), vbLf)
            For f = 0 To UBound(faqLines)
              faqLine = Trim(faqLines(f))
              pipePos = InStr(1, faqLine, "|", 1)
              If faqLine <> "" And pipePos > 0 Then
                faqQuestion = Trim(Left(faqLine, pipePos - 1))
                faqAnswer = Trim(Mid(faqLine, pipePos + 1))
            %>
              <div class="detail-faq-item">
                <h3><%=UygulamaDetailText(faqQuestion)%></h3>
                <p><%=UygulamaDetailText(faqAnswer)%></p>
              </div>
            <%
              End If
            Next
            %>
          </section>
        <% End If %>
      </article>

      <aside class="detail-side">
        <div class="detail-cta">
          <h2>Benzer bir alan icin kesif planlayalim</h2>
          <p>Bu uygulamadaki gibi ihtiyaca gore kapasite, cihaz tipi, dis unite konumu ve servis erisimi birlikte degerlendirilebilir.</p>
          <a href="<%=UygulamaDetailUrl(dCtaUrl)%>" onclick="if(window.analyticsTrack){analyticsTrack('case_study_cta_click', {page: 'uygulama-ornegi', case_slug: '<%=UygulamaDetailAttr(slug)%>', tracking_tag: '<%=UygulamaDetailAttr(dTakip)%>'});}">
            <%=UygulamaDetailText(dCtaMetni)%> <span>&#8594;</span>
          </a>
        </div>

        <% If Trim(dHizmetUrl) <> "" Or Trim(dBolgeUrl) <> "" Then %>
          <div class="detail-side-links">
            <strong>İlgili Sayfalar</strong>
            <% If Trim(dHizmetUrl) <> "" Then %><a href="<%=UygulamaDetailUrl(dHizmetUrl)%>">İlgili hizmet sayfası</a><% End If %>
            <% If Trim(dBolgeUrl) <> "" Then %><a href="<%=UygulamaDetailUrl(dBolgeUrl)%>">İlgili bölge sayfası</a><% End If %>
          </div>
        <% End If %>
      </aside>
    </div>
  </div>
</div>

<!--  #include file='alt.asp' -->

</body>
</html>
