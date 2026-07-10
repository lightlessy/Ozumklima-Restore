<%
Function UygulamaAdminSql(ByVal value)
  UygulamaAdminSql = Replace(CStr(value & ""), "'", "''")
End Function

Function UygulamaAdminHtml(ByVal value)
  UygulamaAdminHtml = Server.HTMLEncode(CStr(value & ""))
End Function

Function UygulamaAdminTextarea(ByVal value)
  UygulamaAdminTextarea = Server.HTMLEncode(CStr(value & ""))
End Function

Function UygulamaAdminNumber(ByVal value, ByVal fallback)
  If IsNumeric(value) Then
    UygulamaAdminNumber = CLng(value)
  Else
    UygulamaAdminNumber = fallback
  End If
End Function

Function UygulamaAdminSlug(ByVal value)
  Dim slugValue
  slugValue = LCase(Trim(CStr(value & "")))
  slugValue = Replace(slugValue, "&", " ve ")
  slugValue = Replace(slugValue, "/", " ")
  slugValue = Replace(slugValue, "\", " ")
  slugValue = Replace(slugValue, ".", " ")
  slugValue = Replace(slugValue, ",", " ")
  slugValue = Replace(slugValue, ":", " ")
  slugValue = Replace(slugValue, ";", " ")
  slugValue = Replace(slugValue, "'", "")
  slugValue = Replace(slugValue, Chr(34), "")
  slugValue = Replace(slugValue, " ", "-")
  Do While InStr(slugValue, "--") > 0
    slugValue = Replace(slugValue, "--", "-")
  Loop
  If slugValue = "" Then slugValue = "uygulama-ornegi"
  UygulamaAdminSlug = slugValue
End Function

Function UygulamaAdminIsLive(ByVal value)
  Dim liveText
  liveText = LCase(Trim(CStr(value & "")))
  UygulamaAdminIsLive = False
  If liveText = "true" Or liveText = "1" Or liveText = "-1" Then UygulamaAdminIsLive = True
End Function

Sub UygulamaAdminSeed(ByVal baslik, ByVal etiket, ByVal problem, ByVal cozum, ByVal uygulama, ByVal sonuc, ByVal ctaMetni, ByVal ctaUrl, ByVal slug, ByVal sira)
  Dim seedSql
  seedSql = "INSERT INTO uygulama_ornekleri ([baslik], [etiket], [problem], [cozum], [uygulama], [sonuc], [cta_metni], [cta_url], [slug], [sira], [yayin], [kayit_tarihi]) VALUES (" & _
            "'" & UygulamaAdminSql(baslik) & "'," & _
            "'" & UygulamaAdminSql(etiket) & "'," & _
            "'" & UygulamaAdminSql(problem) & "'," & _
            "'" & UygulamaAdminSql(cozum) & "'," & _
            "'" & UygulamaAdminSql(uygulama) & "'," & _
            "'" & UygulamaAdminSql(sonuc) & "'," & _
            "'" & UygulamaAdminSql(ctaMetni) & "'," & _
            "'" & UygulamaAdminSql(ctaUrl) & "'," & _
            "'" & UygulamaAdminSql(slug) & "'," & CLng(sira) & ",-1,Now())"
  baglanti.Execute seedSql
End Sub

Sub UygulamaAdminEnsureTable()
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
    Call UygulamaAdminSeed("Antalya Villa Klima Montaji", "Villa / Konut", "Musteri yaz aylarinda salon ve yatak odalarinda yeterli sogutma alamiyordu.", "Alan kesfi yapildi, uygun BTU hesaplandi ve Mitsubishi Electric klima sistemi onerildi.", "Montaj, drenaj hatti ve dis unite konumlandirmasi tamamlandi.", "Musteri daha sessiz, verimli ve stabil bir sogutma deneyimi elde etti.", "Benzer bir cozum icin iletisime gec", "../iletisim.asp", "villa-klima-montaji", 10)
    Call UygulamaAdminSeed("Mitsubishi Electric Klima Bakim Uygulamasi", "Bakim / Servis", "Klimada sogutma performansi dusmus, enerji tuketimi artmis ve hava kalitesi zayiflamisti.", "Filtre, serpantin ve drenaj hatti kontrol edilerek periyodik bakim plani uygulandi.", "Ic ve dis unite temizlik, gaz basinc kontrolu ve calisma testi adimlari tamamlandi.", "Daha dengeli sogutma, dusuk enerji tuketimi ve daha konforlu ic ortam elde edildi.", "Bakim plani al", "../iletisim.asp", "mitsubishi-bakim", 20)
    Call UygulamaAdminSeed("Isyeri / Ofis VRF Sistem Cozumu", "VRF / Ticari", "Farkli ofis bolumlerinde sicaklik dengesi saglanamiyor, enerji maliyetleri yuksek seyrediyordu.", "Bolge bazli ihtiyac analizi yapilarak uygun kapasitede VRF ic ve dis unite kombinasyonu planlandi.", "Boru hatti, otomasyon ayarlari ve oda bazli kontrol senaryolari devreye alindi.", "Ofis genelinde dengeli iklimlendirme ve olculebilir enerji verimliligi elde edildi.", "VRF kesfi talep et", "../iletisim.asp", "ofis-vrf", 30)
  End If
End Sub
%>

<!--#INCLUDE file="menu.asp"-->

<body>

<%
Call UygulamaAdminEnsureTable()

Dim formAction, redirectMsg, recordId, saveRs, deleteId, sortValue, publishValue, slugValue, ctaUrlValue
formAction = Request.Form("form_action")

If formAction = "delete" Then
  deleteId = UygulamaAdminNumber(Request.Form("id"), 0)
  If deleteId > 0 Then
    baglanti.Execute "DELETE FROM uygulama_ornekleri WHERE [id]=" & deleteId
  End If
  Response.Redirect "uygulama-ornekleri.asp?ok=delete"
End If

If formAction = "save" Then
  recordId = UygulamaAdminNumber(Request.Form("id"), 0)
  sortValue = UygulamaAdminNumber(Request.Form("sira"), 0)
  publishValue = 0
  If Request.Form("yayin") = "1" Then publishValue = -1

  slugValue = Trim(Request.Form("slug") & "")
  If slugValue = "" Then slugValue = UygulamaAdminSlug(Request.Form("baslik"))

  ctaUrlValue = Trim(Request.Form("cta_url") & "")
  If ctaUrlValue = "" Then ctaUrlValue = "../iletisim.asp"

  Set saveRs = Server.CreateObject("ADODB.RecordSet")
  If recordId > 0 Then
    saveRs.Open "SELECT * FROM uygulama_ornekleri WHERE [id]=" & recordId, baglanti, 1, 3
  Else
    saveRs.Open "SELECT * FROM uygulama_ornekleri WHERE [id]=0", baglanti, 1, 3
    saveRs.AddNew
  End If

  If recordId = 0 Or Not saveRs.EOF Then
    saveRs("baslik") = Trim(Request.Form("baslik") & "")
    saveRs("etiket") = Trim(Request.Form("etiket") & "")
    saveRs("problem") = Trim(Request.Form("problem") & "")
    saveRs("cozum") = Trim(Request.Form("cozum") & "")
    saveRs("uygulama") = Trim(Request.Form("uygulama") & "")
    saveRs("sonuc") = Trim(Request.Form("sonuc") & "")
    saveRs("cta_metni") = Trim(Request.Form("cta_metni") & "")
    saveRs("cta_url") = ctaUrlValue
    saveRs("slug") = slugValue
    saveRs("sira") = sortValue
    saveRs("yayin") = publishValue
    saveRs.Update
  End If

  If saveRs.State = 1 Then saveRs.Close
  Set saveRs = Nothing

  Response.Redirect "uygulama-ornekleri.asp?ok=save"
End If

Dim editMode, editId, editRs
Dim fBaslik, fEtiket, fProblem, fCozum, fUygulama, fSonuc, fCtaMetni, fCtaUrl, fSlug, fSira, fYayin
editMode = False
editId = UygulamaAdminNumber(Request.QueryString("sid"), 0)
fBaslik = ""
fEtiket = ""
fProblem = ""
fCozum = ""
fUygulama = ""
fSonuc = ""
fCtaMetni = "Iletisime gec"
fCtaUrl = "../iletisim.asp"
fSlug = ""
fSira = 10
fYayin = True

If Request.QueryString("edit") = "yes" And editId > 0 Then
  Set editRs = Server.CreateObject("ADODB.RecordSet")
  editRs.Open "SELECT * FROM uygulama_ornekleri WHERE [id]=" & editId, baglanti, 1, 1
  If Not editRs.EOF Then
    editMode = True
    fBaslik = editRs("baslik") & ""
    fEtiket = editRs("etiket") & ""
    fProblem = editRs("problem") & ""
    fCozum = editRs("cozum") & ""
    fUygulama = editRs("uygulama") & ""
    fSonuc = editRs("sonuc") & ""
    fCtaMetni = editRs("cta_metni") & ""
    fCtaUrl = editRs("cta_url") & ""
    fSlug = editRs("slug") & ""
    fSira = UygulamaAdminNumber(editRs("sira"), 10)
    fYayin = UygulamaAdminIsLive(editRs("yayin"))
  End If
  If editRs.State = 1 Then editRs.Close
  Set editRs = Nothing
End If
%>

<style>
  .case-admin-wrap,
  .case-admin-wrap * {
    box-sizing: border-box;
  }

  .case-admin-wrap {
    max-width: 1070px;
    margin: 0 auto 50px auto;
    padding: 22px 12px 50px 12px;
    font-family: Open Sans, Arial, sans-serif;
    color: #2f2f2f;
  }

  .case-admin-hero {
    border: 1px solid #dcdcdc;
    border-radius: 18px;
    background: linear-gradient(135deg, #ffffff, #f8f1dc);
    padding: 24px;
    margin-bottom: 18px;
    box-shadow: 0 12px 28px rgba(0,0,0,.06);
  }

  .case-admin-hero h1 {
    margin: 0 0 8px 0;
    font-size: 28px;
    color: #252525;
  }

  .case-admin-hero p {
    margin: 0;
    color: #666;
    font-size: 14px;
    line-height: 1.6em;
  }

  .case-admin-alert {
    margin-bottom: 16px;
    padding: 12px 14px;
    border-radius: 12px;
    background: #eef8e8;
    border: 1px solid #cfe7c4;
    color: #315b20;
    font-size: 13px;
    font-weight: 700;
  }

  .case-admin-card {
    border: 1px solid #dedede;
    border-radius: 16px;
    background: #ffffff;
    margin-bottom: 18px;
    box-shadow: 0 10px 24px rgba(0,0,0,.045);
    overflow: hidden;
  }

  .case-admin-card-head {
    display: flex;
    align-items: center;
    justify-content: space-between;
    gap: 14px;
    padding: 16px 18px;
    background: #3c3c3c;
    color: #ffffff;
  }

  .case-admin-card-head strong {
    color: #f7bb09;
    font-size: 17px;
  }

  .case-admin-card-body {
    padding: 18px;
  }

  .case-admin-grid {
    display: grid;
    grid-template-columns: repeat(2, minmax(0, 1fr));
    gap: 14px;
  }

  .case-admin-field {
    margin-bottom: 14px;
  }

  .case-admin-field label {
    display: block;
    margin-bottom: 6px;
    color: #333;
    font-size: 13px;
    font-weight: 800;
  }

  .case-admin-field input[type=text],
  .case-admin-field textarea {
    width: 100%;
    border: 1px solid #d7d7d7;
    border-radius: 10px;
    background: #fbfbfb;
    padding: 10px 11px;
    color: #2f2f2f;
    font-size: 14px;
    outline: none;
  }

  .case-admin-field textarea {
    min-height: 92px;
    resize: vertical;
    line-height: 1.55em;
  }

  .case-admin-field input[type=text]:focus,
  .case-admin-field textarea:focus {
    border-color: #f7bb09;
    background: #ffffff;
    box-shadow: 0 0 0 3px rgba(247,187,9,.18);
  }

  .case-admin-check {
    display: inline-flex;
    align-items: center;
    gap: 8px;
    padding: 10px 12px;
    border-radius: 10px;
    background: #f7f7f7;
    font-size: 13px;
    font-weight: 800;
  }

  .case-admin-actions {
    display: flex;
    flex-wrap: wrap;
    align-items: center;
    gap: 10px;
    margin-top: 4px;
  }

  .case-admin-button,
  .case-admin-link,
  .case-admin-danger {
    display: inline-flex;
    align-items: center;
    justify-content: center;
    border: 0;
    border-radius: 10px;
    padding: 10px 16px;
    font-size: 14px;
    font-weight: 800;
    cursor: pointer;
    text-decoration: none;
  }

  .case-admin-button {
    background: #5e8f2c;
    color: #ffffff;
  }

  .case-admin-link {
    background: #efefef;
    color: #333333;
  }

  .case-admin-danger {
    background: #ba2d2d;
    color: #ffffff;
  }

  .case-admin-table-wrap {
    overflow-x: auto;
  }

  .case-admin-table {
    width: 100%;
    border-collapse: collapse;
    min-width: 850px;
    font-size: 13px;
  }

  .case-admin-table th {
    background: #f3f0e4;
    color: #333;
    padding: 11px 10px;
    text-align: left;
    border-bottom: 1px solid #dedede;
  }

  .case-admin-table td {
    padding: 12px 10px;
    border-bottom: 1px solid #ededed;
    vertical-align: top;
  }

  .case-admin-status {
    display: inline-block;
    border-radius: 999px;
    padding: 5px 9px;
    font-size: 12px;
    font-weight: 800;
    background: #ededed;
    color: #555;
  }

  .case-admin-status.live {
    background: #e7f5de;
    color: #315b20;
  }

  @media screen and (max-width: 760px) {
    .case-admin-grid {
      grid-template-columns: 1fr;
    }

    .case-admin-card-head {
      display: block;
    }
  }
</style>

<div class="baslik">Uygulama Ornekleri</div>

<div class="case-admin-wrap">
  <div class="case-admin-hero">
    <h1>Uygulama ornekleri paneli</h1>
    <p>Bu sayfadan sitedeki uygulama orneklerini ekleyebilir, duzenleyebilir, silebilir ve yayindan kaldirabilirsiniz. Ilk acilista tablo yoksa otomatik olusturulur ve mevcut statik ornekler veritabanina tasinir.</p>
  </div>

  <% If Request.QueryString("ok") = "save" Then %>
    <div class="case-admin-alert">Kayit basariyla kaydedildi.</div>
  <% ElseIf Request.QueryString("ok") = "delete" Then %>
    <div class="case-admin-alert">Kayit silindi.</div>
  <% End If %>

  <div class="case-admin-card">
    <div class="case-admin-card-head">
      <strong><% If editMode Then %>Kaydi duzenle<% Else %>Yeni uygulama ornegi ekle<% End If %></strong>
      <span>Problem &rarr; Cozum &rarr; Uygulama &rarr; Sonuc akisi</span>
    </div>

    <div class="case-admin-card-body">
      <form method="post" action="uygulama-ornekleri.asp">
        <input type="hidden" name="form_action" value="save">
        <input type="hidden" name="id" value="<%=editId%>">

        <div class="case-admin-grid">
          <div class="case-admin-field">
            <label>Baslik</label>
            <input type="text" name="baslik" value="<%=UygulamaAdminHtml(fBaslik)%>" maxlength="255" required>
          </div>

          <div class="case-admin-field">
            <label>Etiket</label>
            <input type="text" name="etiket" value="<%=UygulamaAdminHtml(fEtiket)%>" maxlength="100" placeholder="Villa / VRF / Bakim">
          </div>
        </div>

        <div class="case-admin-grid">
          <div class="case-admin-field">
            <label>Problem</label>
            <textarea name="problem"><%=UygulamaAdminTextarea(fProblem)%></textarea>
          </div>

          <div class="case-admin-field">
            <label>Cozum</label>
            <textarea name="cozum"><%=UygulamaAdminTextarea(fCozum)%></textarea>
          </div>
        </div>

        <div class="case-admin-grid">
          <div class="case-admin-field">
            <label>Uygulama</label>
            <textarea name="uygulama"><%=UygulamaAdminTextarea(fUygulama)%></textarea>
          </div>

          <div class="case-admin-field">
            <label>Sonuc</label>
            <textarea name="sonuc"><%=UygulamaAdminTextarea(fSonuc)%></textarea>
          </div>
        </div>

        <div class="case-admin-grid">
          <div class="case-admin-field">
            <label>CTA metni</label>
            <input type="text" name="cta_metni" value="<%=UygulamaAdminHtml(fCtaMetni)%>" maxlength="255">
          </div>

          <div class="case-admin-field">
            <label>CTA linki</label>
            <input type="text" name="cta_url" value="<%=UygulamaAdminHtml(fCtaUrl)%>" maxlength="255" placeholder="../iletisim.asp">
          </div>
        </div>

        <div class="case-admin-grid">
          <div class="case-admin-field">
            <label>Slug / takip etiketi</label>
            <input type="text" name="slug" value="<%=UygulamaAdminHtml(fSlug)%>" maxlength="120" placeholder="Bos kalirsa basliktan uretilir">
          </div>

          <div class="case-admin-field">
            <label>Siralama</label>
            <input type="text" name="sira" value="<%=UygulamaAdminHtml(fSira)%>" maxlength="6">
          </div>
        </div>

        <div class="case-admin-field">
          <label class="case-admin-check"><input type="checkbox" name="yayin" value="1" <% If fYayin Then Response.Write "checked" %>> Yayinda goster</label>
        </div>

        <div class="case-admin-actions">
          <button type="submit" class="case-admin-button"><% If editMode Then %>Guncelle<% Else %>Kaydet<% End If %></button>
          <% If editMode Then %><a href="uygulama-ornekleri.asp" class="case-admin-link">Yeni kayit moduna don</a><% End If %>
          <a href="../uygulama-ornekleri.asp" target="_blank" class="case-admin-link">Sitede gor</a>
        </div>
      </form>
    </div>
  </div>

  <div class="case-admin-card">
    <div class="case-admin-card-head">
      <strong>Kayit listesi</strong>
      <span>Sira numarasina gore yayinda gorunur</span>
    </div>

    <div class="case-admin-card-body case-admin-table-wrap">
      <%
      Dim listRs
      Set listRs = Server.CreateObject("ADODB.RecordSet")
      listRs.Open "SELECT * FROM uygulama_ornekleri ORDER BY [sira] ASC, [id] ASC", baglanti, 1, 1
      %>

      <table class="case-admin-table">
        <tr>
          <th width="55">Sira</th>
          <th>Baslik</th>
          <th width="130">Etiket</th>
          <th width="110">Durum</th>
          <th width="95">Duzenle</th>
          <th width="75">Sil</th>
        </tr>

        <% If listRs.EOF Then %>
          <tr><td colspan="6">Henuz kayit yok.</td></tr>
        <% Else %>
          <%
          Do While Not listRs.EOF
          %>
            <tr>
              <td><%=UygulamaAdminHtml(listRs("sira"))%></td>
              <td><b><%=UygulamaAdminHtml(listRs("baslik"))%></b><br><span style="color:#777"><%=UygulamaAdminHtml(listRs("slug"))%></span></td>
              <td><%=UygulamaAdminHtml(listRs("etiket"))%></td>
              <td>
                <% If UygulamaAdminIsLive(listRs("yayin")) Then %>
                  <span class="case-admin-status live">Yayinda</span>
                <% Else %>
                  <span class="case-admin-status">Pasif</span>
                <% End If %>
              </td>
              <td><a class="case-admin-link" href="uygulama-ornekleri.asp?edit=yes&sid=<%=listRs("id")%>">Duzenle</a></td>
              <td>
                <form method="post" action="uygulama-ornekleri.asp" style="margin:0">
                  <input type="hidden" name="form_action" value="delete">
                  <input type="hidden" name="id" value="<%=listRs("id")%>">
                  <button type="submit" class="case-admin-danger" onclick="return confirm('Bu uygulama ornegini silmek istiyor musunuz?')">Sil</button>
                </form>
              </td>
            </tr>
          <%
            listRs.MoveNext
          Loop
          %>
        <% End If %>
      </table>

      <%
      If listRs.State = 1 Then listRs.Close
      Set listRs = Nothing
      %>
    </div>
  </div>
</div>

</body>