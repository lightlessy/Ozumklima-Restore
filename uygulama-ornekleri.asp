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

Function UygulamaCaseUrl(ByVal value)
  Dim urlValue, lowerUrl
  urlValue = Trim(CStr(value & ""))
  If urlValue = "" Then urlValue = "iletisim.asp"

  lowerUrl = LCase(urlValue)
  If InStr(1, lowerUrl, "javascript:", 1) > 0 Then urlValue = "iletisim.asp"

  UygulamaCaseUrl = Server.HTMLEncode(urlValue)
End Function

Sub UygulamaCaseSeed(ByVal baslik, ByVal etiket, ByVal problem, ByVal cozum, ByVal uygulama, ByVal sonuc, ByVal ctaMetni, ByVal ctaUrl, ByVal slug, ByVal sira)
  Dim seedSql
  seedSql = "INSERT INTO uygulama_ornekleri ([baslik], [etiket], [problem], [cozum], [uygulama], [sonuc], [cta_metni], [cta_url], [slug], [sira], [yayin], [kayit_tarihi]) VALUES (" & _
            "'" & UygulamaCaseSql(baslik) & "'," & _
            "'" & UygulamaCaseSql(etiket) & "'," & _
            "'" & UygulamaCaseSql(problem) & "'," & _
            "'" & UygulamaCaseSql(cozum) & "'," & _
            "'" & UygulamaCaseSql(uygulama) & "'," & _
            "'" & UygulamaCaseSql(sonuc) & "'," & _
            "'" & UygulamaCaseSql(ctaMetni) & "'," & _
            "'" & UygulamaCaseSql(ctaUrl) & "'," & _
            "'" & UygulamaCaseSql(slug) & "'," & CLng(sira) & ",-1,Now())"
  baglanti.Execute seedSql
End Sub

Sub UygulamaCaseEnsureTable()
  Dim testRs, toplamKayit
  toplamKayit = 0

  On Error Resume Next
  Set testRs = baglanti.Execute("SELECT TOP 1 [id] FROM uygulama_ornekleri")
  If Err.Number <> 0 Then
    Err.Clear
    baglanti.Execute "CREATE TABLE uygulama_ornekleri ([id] AUTOINCREMENT PRIMARY KEY, [baslik] TEXT(255), [etiket] TEXT(100), [problem] MEMO, [cozum] MEMO, [uygulama] MEMO, [sonuc] MEMO, [cta_metni] TEXT(255), [cta_url] TEXT(255), [slug] TEXT(120), [sira] INTEGER, [yayin] YESNO, [kayit_tarihi] DATETIME)"
  End If
  If Not testRs Is Nothing Then
    If testRs.State = 1 Then testRs.Close
    Set testRs = Nothing
  End If
  On Error GoTo 0

  On Error Resume Next
  Set testRs = baglanti.Execute("SELECT COUNT(*) AS toplam FROM uygulama_ornekleri")
  If Err.Number = 0 Then toplamKayit = CLng(testRs("toplam"))
  If Not testRs Is Nothing Then
    If testRs.State = 1 Then testRs.Close
    Set testRs = Nothing
  End If
  On Error GoTo 0

  If toplamKayit = 0 Then
    Call UygulamaCaseSeed("Antalya Villa Klima Montaji", "Villa / Konut", "Musteri yaz aylarinda salon ve yatak odalarinda yeterli sogutma alamiyordu.", "Alan kesfi yapildi, uygun BTU hesaplandi ve Mitsubishi Electric klima sistemi onerildi.", "Montaj, drenaj hatti ve dis unite konumlandirmasi tamamlandi.", "Musteri daha sessiz, verimli ve stabil bir sogutma deneyimi elde etti.", "Benzer bir cozum icin iletisime gec", "iletisim.asp", "villa-klima-montaji", 10)
    Call UygulamaCaseSeed("Mitsubishi Electric Klima Bakim Uygulamasi", "Bakim / Servis", "Klimada sogutma performansi dusmus, enerji tuketimi artmis ve hava kalitesi zayiflamisti.", "Filtre, serpantin ve drenaj hatti kontrol edilerek periyodik bakim plani uygulandi.", "Ic ve dis unite temizlik, gaz basinc kontrolu ve calisma testi adimlari tamamlandi.", "Daha dengeli sogutma, dusuk enerji tuketimi ve daha konforlu ic ortam elde edildi.", "Bakim plani al", "iletisim.asp", "mitsubishi-bakim", 20)
    Call UygulamaCaseSeed("Isyeri / Ofis VRF Sistem Cozumu", "VRF / Ticari", "Farkli ofis bolumlerinde sicaklik dengesi saglanamiyor, enerji maliyetleri yuksek seyrediyordu.", "Bolge bazli ihtiyac analizi yapilarak uygun kapasitede VRF ic ve dis unite kombinasyonu planlandi.", "Boru hatti, otomasyon ayarlari ve oda bazli kontrol senaryolari devreye alindi.", "Ofis genelinde dengeli iklimlendirme ve olculebilir enerji verimliligi elde edildi.", "VRF kesfi talep et", "iletisim.asp", "ofis-vrf", 30)
  End If
End Sub
%>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html>
<head>

<META HTTP-EQUIV="Content-Type" content="text/html; charset=windows-1254">
<meta name="viewport" content="width=device-width, initial-scale=1">

<title>Uygulama Ornekleri ve Case Study | Ozum Klima Antalya</title>
<meta name="description" content="Ozum Klima'nin Antalya'da tamamladigi villa klima montaji, Mitsubishi Electric bakim ve ofis VRF uygulama orneklerini inceleyin." />
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
    padding: 0 0 54px 0;
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
    border-radius: 20px;
    background: linear-gradient(135deg, #ffffff 0%, #f7f1df 100%);
    padding: 34px;
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
    max-width: 760px;
    font-size: 2.55em;
    line-height: 1.08em;
    color: #252525;
    font-weight: 800;
  }

  .case-subtitle {
    position: relative;
    z-index: 1;
    max-width: 720px;
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
    gap: 18px;
  }

  .case-card {
    display: flex;
    flex-direction: column;
    min-height: 100%;
    border: 1px solid #e6e6e6;
    border-radius: 18px;
    background: #ffffff;
    padding: 22px;
    box-shadow: 0 14px 34px rgba(0,0,0,.055);
    transition: transform .18s ease, box-shadow .18s ease, border-color .18s ease;
  }

  .case-card:hover {
    transform: translateY(-3px);
    border-color: #f7bb09;
    box-shadow: 0 18px 42px rgba(0,0,0,.085);
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
    margin: 0 0 16px 0;
    font-size: 1.36em;
    line-height: 1.28em;
    color: #232323;
    font-weight: 800;
  }

  .case-field {
    margin-bottom: 13px;
    padding-left: 13px;
    border-left: 3px solid #f0d27b;
    font-size: 1em;
    line-height: 1.55em;
    color: #424242;
  }

  .case-field b {
    display: block;
    margin-bottom: 2px;
    color: #242424;
    font-size: .88em;
    letter-spacing: .04em;
    text-transform: uppercase;
  }

  .case-cta {
    margin-top: auto;
    padding-top: 8px;
  }

  .case-cta a {
    display: inline-flex;
    align-items: center;
    justify-content: center;
    gap: 8px;
    width: 100%;
    min-height: 44px;
    padding: 11px 15px;
    border: 1px solid #2f2f2f;
    border-radius: 12px;
    background: #3c3c3c;
    color: #ffffff;
    font-weight: 800;
    text-decoration: none;
    transition: background .18s ease, color .18s ease, border-color .18s ease;
  }

  .case-cta a:hover {
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
      <p class="case-subtitle">Ozum Klima'nin Antalya'da hayata gecirdigi gercek saha uygulamalarindan secili ozetler. Problem, cozum, uygulama ve sonuc adimlarini tek ekranda inceleyin.</p>

      <div class="case-metrics">
        <div class="case-metric"><strong>Kesif</strong><span>Ihtiyaca gore kapasite analizi</span></div>
        <div class="case-metric"><strong>Uygulama</strong><span>Montaj, bakim ve VRF surecleri</span></div>
        <div class="case-metric"><strong>Sonuc</strong><span>Konfor ve verimlilik odakli cikti</span></div>
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
        %>
          <article class="case-card">
            <% If Trim(uygulamaRs("etiket") & "") <> "" Then %>
              <div class="case-badge"><%=UygulamaCaseText(uygulamaRs("etiket"))%></div>
            <% End If %>

            <h2><%=UygulamaCaseText(uygulamaRs("baslik"))%></h2>

            <div class="case-field"><b>Problem</b><%=UygulamaCaseText(uygulamaRs("problem"))%></div>
            <div class="case-field"><b>Cozum</b><%=UygulamaCaseText(uygulamaRs("cozum"))%></div>
            <div class="case-field"><b>Uygulama</b><%=UygulamaCaseText(uygulamaRs("uygulama"))%></div>
            <div class="case-field"><b>Sonuc</b><%=UygulamaCaseText(uygulamaRs("sonuc"))%></div>

            <div class="case-cta">
              <a href="<%=UygulamaCaseUrl(uygulamaRs("cta_url"))%>" onclick="if(window.analyticsTrack){analyticsTrack('case_study_cta_click', {page: 'uygulama-ornekleri', case_slug: '<%=UygulamaCaseJs(uygulamaRs("slug"))%>'});}">
                <%=UygulamaCaseText(uygulamaRs("cta_metni"))%> <span>&#8594;</span>
              </a>
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