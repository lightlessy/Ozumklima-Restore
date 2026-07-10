<% Response.Buffer = True %>
<!--#INCLUDE file="dbase.asp"-->
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

Function UygulamaAdminValue(ByRef rs, ByVal fieldName)
  On Error Resume Next
  UygulamaAdminValue = rs(fieldName) & ""
  If Err.Number <> 0 Then
    UygulamaAdminValue = ""
    Err.Clear
  End If
  On Error GoTo 0
End Function

Sub UygulamaAdminSetField(ByRef rs, ByVal fieldName, ByVal fieldValue)
  On Error Resume Next
  rs(fieldName) = fieldValue
  If Err.Number <> 0 Then Err.Clear
  On Error GoTo 0
End Sub

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
  slugValue = Replace(slugValue, "(", " ")
  slugValue = Replace(slugValue, ")", " ")
  slugValue = Replace(slugValue, " ", "-")
  Do While InStr(slugValue, "--") > 0
    slugValue = Replace(slugValue, "--", "-")
  Loop
  If Left(slugValue, 1) = "-" Then slugValue = Mid(slugValue, 2)
  If Right(slugValue, 1) = "-" Then slugValue = Left(slugValue, Len(slugValue) - 1)
  If slugValue = "" Then slugValue = "uygulama-ornegi"
  UygulamaAdminSlug = slugValue
End Function

Function UygulamaAdminIsLive(ByVal value)
  Dim liveText
  liveText = LCase(Trim(CStr(value & "")))
  UygulamaAdminIsLive = False
  If liveText = "true" Or liveText = "1" Or liveText = "-1" Then UygulamaAdminIsLive = True
End Function

Function UygulamaAdminSafeTarget(ByVal value)
  Dim targetValue
  targetValue = LCase(Trim(CStr(value & "")))
  Select Case targetValue
    Case "cover_image", "galeri_gorseller", "oncesi_image", "sonrasi_image"
      UygulamaAdminSafeTarget = targetValue
    Case Else
      UygulamaAdminSafeTarget = "cover_image"
  End Select
End Function

Function UygulamaAdminMediaPreviewSrc(ByVal value)
  Dim imageValue, lowerImage
  imageValue = Trim(CStr(value & ""))
  lowerImage = LCase(imageValue)
  If imageValue = "" Then
    UygulamaAdminMediaPreviewSrc = ""
    Exit Function
  End If
  If InStr(1, lowerImage, "javascript:", 1) > 0 Then
    UygulamaAdminMediaPreviewSrc = ""
    Exit Function
  End If
  If Left(lowerImage, 4) = "http" Or Left(imageValue, 1) = "/" Or Left(imageValue, 3) = "../" Then
    UygulamaAdminMediaPreviewSrc = imageValue
  Else
    UygulamaAdminMediaPreviewSrc = "../" & imageValue
  End If
End Function

Function UygulamaAdminFileExt(ByVal value)
  Dim fileValue, dotPos
  fileValue = CStr(value & "")
  dotPos = InStrRev(fileValue, ".")
  If dotPos > 0 Then
    UygulamaAdminFileExt = LCase(Mid(fileValue, dotPos + 1))
  Else
    UygulamaAdminFileExt = ""
  End If
End Function

Function UygulamaAdminIsImagePath(ByVal value)
  Dim ext
  ext = UygulamaAdminFileExt(value)
  UygulamaAdminIsImagePath = False
  If InStr(1, "|jpg|jpeg|png|gif|webp|", "|" & ext & "|", 1) > 0 Then UygulamaAdminIsImagePath = True
End Function

Sub UygulamaAdminAddColumn(ByVal columnSql)
  On Error Resume Next
  baglanti.Execute "ALTER TABLE uygulama_ornekleri ADD COLUMN " & columnSql
  Err.Clear
  On Error GoTo 0
End Sub

Sub UygulamaAdminAddMediaColumn(ByVal columnSql)
  On Error Resume Next
  baglanti.Execute "ALTER TABLE uygulama_medya ADD COLUMN " & columnSql
  Err.Clear
  On Error GoTo 0
End Sub

Sub UygulamaAdminSeed(ByVal baslik, ByVal etiket, ByVal kisaOzet, ByVal lokasyon, ByVal projeTipi, ByVal hizmetTipi, ByVal sistemMarka, ByVal projeSuresi, ByVal problem, ByVal kesif, ByVal cozum, ByVal uygulama, ByVal sonuc, ByVal ctaMetni, ByVal ctaUrl, ByVal slug, ByVal takipEtiketi, ByVal sira)
  Dim seedSql
  seedSql = "INSERT INTO uygulama_ornekleri ([baslik], [etiket], [kisa_ozet], [lokasyon], [proje_tipi], [hizmet_tipi], [sistem_marka], [proje_suresi], [problem], [kesif_degerlendirme], [cozum], [uygulama], [sonuc], [cta_metni], [cta_url], [slug], [takip_etiketi], [sira], [yayin], [kayit_tarihi]) VALUES (" & _
            "'" & UygulamaAdminSql(baslik) & "'," & _
            "'" & UygulamaAdminSql(etiket) & "'," & _
            "'" & UygulamaAdminSql(kisaOzet) & "'," & _
            "'" & UygulamaAdminSql(lokasyon) & "'," & _
            "'" & UygulamaAdminSql(projeTipi) & "'," & _
            "'" & UygulamaAdminSql(hizmetTipi) & "'," & _
            "'" & UygulamaAdminSql(sistemMarka) & "'," & _
            "'" & UygulamaAdminSql(projeSuresi) & "'," & _
            "'" & UygulamaAdminSql(problem) & "'," & _
            "'" & UygulamaAdminSql(kesif) & "'," & _
            "'" & UygulamaAdminSql(cozum) & "'," & _
            "'" & UygulamaAdminSql(uygulama) & "'," & _
            "'" & UygulamaAdminSql(sonuc) & "'," & _
            "'" & UygulamaAdminSql(ctaMetni) & "'," & _
            "'" & UygulamaAdminSql(ctaUrl) & "'," & _
            "'" & UygulamaAdminSql(slug) & "'," & _
            "'" & UygulamaAdminSql(takipEtiketi) & "'," & CLng(sira) & ",-1,Now())"
  baglanti.Execute seedSql
End Sub

Sub UygulamaAdminEnsureMediaTable()
  Dim testRs
  On Error Resume Next
  Set testRs = baglanti.Execute("SELECT TOP 1 [id] FROM uygulama_medya")
  If Err.Number <> 0 Then
    Err.Clear
    baglanti.Execute "CREATE TABLE uygulama_medya ([id] AUTOINCREMENT PRIMARY KEY, [dosya_yolu] TEXT(255), [alt_text] TEXT(255), [kaynak] TEXT(60), [kayit_tarihi] DATETIME)"
  End If
  If Not testRs Is Nothing Then
    If testRs.State = 1 Then testRs.Close
    Set testRs = Nothing
  End If
  On Error GoTo 0

  Call UygulamaAdminAddMediaColumn("[dosya_yolu] TEXT(255)")
  Call UygulamaAdminAddMediaColumn("[alt_text] TEXT(255)")
  Call UygulamaAdminAddMediaColumn("[kaynak] TEXT(60)")
  Call UygulamaAdminAddMediaColumn("[kayit_tarihi] DATETIME")
End Sub

Sub UygulamaAdminRegisterMedia(ByVal mediaPath, ByVal altText, ByVal sourceText)
  If Trim(CStr(mediaPath & "")) = "" Then Exit Sub
  Call UygulamaAdminEnsureMediaTable()
  On Error Resume Next
  baglanti.Execute "INSERT INTO uygulama_medya ([dosya_yolu], [alt_text], [kaynak], [kayit_tarihi]) VALUES ('" & UygulamaAdminSql(mediaPath) & "','" & UygulamaAdminSql(altText) & "','" & UygulamaAdminSql(sourceText) & "',Now())"
  Err.Clear
  On Error GoTo 0
End Sub

Sub UygulamaAdminEnsureTable()
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

  Call UygulamaAdminAddColumn("[baslik] TEXT(255)")
  Call UygulamaAdminAddColumn("[etiket] TEXT(100)")
  Call UygulamaAdminAddColumn("[kisa_ozet] MEMO")
  Call UygulamaAdminAddColumn("[lokasyon] TEXT(150)")
  Call UygulamaAdminAddColumn("[proje_tipi] TEXT(120)")
  Call UygulamaAdminAddColumn("[hizmet_tipi] TEXT(150)")
  Call UygulamaAdminAddColumn("[sistem_marka] TEXT(180)")
  Call UygulamaAdminAddColumn("[proje_suresi] TEXT(80)")
  Call UygulamaAdminAddColumn("[problem] MEMO")
  Call UygulamaAdminAddColumn("[kesif_degerlendirme] MEMO")
  Call UygulamaAdminAddColumn("[cozum] MEMO")
  Call UygulamaAdminAddColumn("[uygulama] MEMO")
  Call UygulamaAdminAddColumn("[sonuc] MEMO")
  Call UygulamaAdminAddColumn("[cover_image] TEXT(255)")
  Call UygulamaAdminAddColumn("[cover_alt] TEXT(255)")
  Call UygulamaAdminAddColumn("[galeri_gorseller] MEMO")
  Call UygulamaAdminAddColumn("[galeri_altlar] MEMO")
  Call UygulamaAdminAddColumn("[video_url] TEXT(255)")
  Call UygulamaAdminAddColumn("[oncesi_image] TEXT(255)")
  Call UygulamaAdminAddColumn("[oncesi_alt] TEXT(255)")
  Call UygulamaAdminAddColumn("[sonrasi_image] TEXT(255)")
  Call UygulamaAdminAddColumn("[sonrasi_alt] TEXT(255)")
  Call UygulamaAdminAddColumn("[musteri_yorumu] MEMO")
  Call UygulamaAdminAddColumn("[sss] MEMO")
  Call UygulamaAdminAddColumn("[ilgili_hizmet_url] TEXT(255)")
  Call UygulamaAdminAddColumn("[ilgili_bolge_url] TEXT(255)")
  Call UygulamaAdminAddColumn("[cta_metni] TEXT(255)")
  Call UygulamaAdminAddColumn("[cta_url] TEXT(255)")
  Call UygulamaAdminAddColumn("[slug] TEXT(120)")
  Call UygulamaAdminAddColumn("[takip_etiketi] TEXT(120)")
  Call UygulamaAdminAddColumn("[seo_title] TEXT(255)")
  Call UygulamaAdminAddColumn("[meta_description] MEMO")
  Call UygulamaAdminAddColumn("[sira] INTEGER")
  Call UygulamaAdminAddColumn("[yayin] YESNO")
  Call UygulamaAdminAddColumn("[kayit_tarihi] DATETIME")

  On Error Resume Next
  Set testRs = baglanti.Execute("SELECT COUNT(*) AS toplam FROM uygulama_ornekleri")
  If Err.Number = 0 Then toplamKayit = CLng(testRs("toplam"))
  If Not testRs Is Nothing Then
    If testRs.State = 1 Then testRs.Close
    Set testRs = Nothing
  End If
  On Error GoTo 0

  If toplamKayit = 0 Then
    Call UygulamaAdminSeed("Antalya Villa Klima Montaji", "Villa / Konut", "Konyaalti bolgesindeki villada salon ve yatak odalari icin sessiz, dengeli ve estetik klima cozumunun nasil planlandigini inceleyin.", "Konyaalti / Antalya", "Villa", "Klima montaji", "Mitsubishi Electric multi split", "1 gun", "Musteri yaz aylarinda salon ve yatak odalarinda yeterli sogutma alamiyor, dis unite konumunun goruntu ve ses acisindan sorun yaratmamasini istiyordu.", "Kesifte oda metrekareleri, cephe yonu, cam alani, drenaj hattinin gececegi guzergah, elektrik altyapisi ve dis unite servis erisimi kontrol edildi.", "Alan ihtiyacina uygun kapasite hesaplandi ve birden fazla ic uniteyi dengeli calistirabilecek Mitsubishi Electric cozum onerildi.", "Ic unite yerleri musteriyle netlestirildi, boru ve drenaj hatlari gizli guzergahlarla tasindi, dis unite bakim erisimi korunacak sekilde konumlandirildi.", "Musteri daha sessiz, daha stabil ve villa mimarisine daha uyumlu bir sogutma deneyimi elde etti.", "Benzer bir cozum icin kesif talep et", "iletisim.asp", "konyaalti-villa-klima-montaji", "case_konyaalti_villa_klima", 10)
  End If
End Sub

Sub UygulamaAdminAddMediaCandidate(ByRef dict, ByVal rawValue)
  Dim cleanValue, lines, i, onePath, lowerPath
  cleanValue = Replace(Replace(CStr(rawValue & ""), vbCrLf, vbLf), vbCr, vbLf)
  lines = Split(cleanValue, vbLf)
  For i = 0 To UBound(lines)
    onePath = Trim(lines(i))
    lowerPath = LCase(onePath)
    If onePath <> "" And InStr(1, lowerPath, "javascript:", 1) = 0 And UygulamaAdminIsImagePath(onePath) Then
      If Not dict.Exists(onePath) Then dict.Add onePath, onePath
    End If
  Next
End Sub

Sub UygulamaAdminRenderMediaLibrary()
  Dim dict, mediaRs, caseRs, fso, folderPath, folderObj, fileObj, mediaPath, mediaCount
  Set dict = Server.CreateObject("Scripting.Dictionary")
  dict.CompareMode = 1

  On Error Resume Next
  Set mediaRs = baglanti.Execute("SELECT [dosya_yolu] FROM uygulama_medya ORDER BY [id] DESC")
  If Err.Number = 0 Then
    Do While Not mediaRs.EOF
      Call UygulamaAdminAddMediaCandidate(dict, UygulamaAdminValue(mediaRs, "dosya_yolu"))
      mediaRs.MoveNext
    Loop
  End If
  Err.Clear
  If Not mediaRs Is Nothing Then
    If mediaRs.State = 1 Then mediaRs.Close
    Set mediaRs = Nothing
  End If

  Set caseRs = baglanti.Execute("SELECT [cover_image], [galeri_gorseller], [oncesi_image], [sonrasi_image] FROM uygulama_ornekleri")
  If Err.Number = 0 Then
    Do While Not caseRs.EOF
      Call UygulamaAdminAddMediaCandidate(dict, UygulamaAdminValue(caseRs, "cover_image"))
      Call UygulamaAdminAddMediaCandidate(dict, UygulamaAdminValue(caseRs, "galeri_gorseller"))
      Call UygulamaAdminAddMediaCandidate(dict, UygulamaAdminValue(caseRs, "oncesi_image"))
      Call UygulamaAdminAddMediaCandidate(dict, UygulamaAdminValue(caseRs, "sonrasi_image"))
      caseRs.MoveNext
    Loop
  End If
  Err.Clear
  If Not caseRs Is Nothing Then
    If caseRs.State = 1 Then caseRs.Close
    Set caseRs = Nothing
  End If

  Set fso = CreateObject("Scripting.FileSystemObject")
  folderPath = Server.MapPath("../images/case")
  If Not fso.FolderExists(folderPath) Then fso.CreateFolder(folderPath)
  If fso.FolderExists(folderPath) Then
    Set folderObj = fso.GetFolder(folderPath)
    For Each fileObj In folderObj.Files
      If UygulamaAdminIsImagePath(fileObj.Name) Then Call UygulamaAdminAddMediaCandidate(dict, "images/case/" & fileObj.Name)
    Next
  End If
  Set folderObj = Nothing
  Set fso = Nothing
  On Error GoTo 0

  If dict.Count = 0 Then
    Response.Write "<div class='case-media-empty'>Henuz medya yok. Once cihazdan bir gorsel yukleyin.</div>"
  Else
    mediaCount = 0
    Response.Write "<div class='case-media-library'>"
    For Each mediaPath In dict.Keys
      mediaCount = mediaCount + 1
      If mediaCount <= 80 Then
        Response.Write "<button type='button' class='case-media-item' data-media-path='" & UygulamaAdminHtml(mediaPath) & "' onclick='casePickMedia(this.getAttribute(""data-media-path""))'>"
        Response.Write "<img src='" & UygulamaAdminHtml(UygulamaAdminMediaPreviewSrc(mediaPath)) & "' alt=''>"
        Response.Write "<span>" & UygulamaAdminHtml(mediaPath) & "</span>"
        Response.Write "</button>"
      End If
    Next
    Response.Write "</div>"
  End If
  Set dict = Nothing
End Sub

Sub UygulamaAdminHandleMediaUpload()
  Dim Upload, File1, ObjFso, saveFolder, savePath, oldPath, newPath, originalName, ext, newName, stamp, publicPath
  Dim targetValue, altText, recordId, redirectUrl, uploadError

  If Request.Cookies("ozum")("wpAfLogin") <> "true" Then
    Response.Redirect "index.asp"
    Response.End
  End If

  uploadError = ""
  saveFolder = "../images/case/"
  savePath = Server.MapPath(saveFolder)

  Set ObjFso = CreateObject("Scripting.FileSystemObject")
  On Error Resume Next
  If Not ObjFso.FolderExists(savePath) Then ObjFso.CreateFolder(savePath)
  If Err.Number <> 0 Then
    uploadError = "folder"
    Err.Clear
  End If
  On Error GoTo 0

  If uploadError = "" Then
    On Error Resume Next
    Set Upload = Server.CreateObject("Persits.Upload")
    Upload.Save savePath
    If Err.Number <> 0 Then
      uploadError = "upload"
      Err.Clear
    End If
    On Error GoTo 0
  End If

  If uploadError = "" Then
    On Error Resume Next
    Set File1 = Upload.Files("resim")
    If Err.Number <> 0 Then
      uploadError = "file"
      Err.Clear
    End If
    On Error GoTo 0
  End If

  If uploadError = "" Then
    If File1 Is Nothing Then uploadError = "file"
  End If

  If uploadError = "" Then
    originalName = File1.FileName
    ext = UygulamaAdminFileExt(originalName)
    targetValue = UygulamaAdminSafeTarget(Upload.Form("media_target"))
    altText = Trim(CStr(Upload.Form("alt_text") & ""))
    recordId = UygulamaAdminNumber(Upload.Form("id"), 0)

    oldPath = savePath & "\" & originalName
    If Not UygulamaAdminIsImagePath(originalName) Then
      On Error Resume Next
      If ObjFso.FileExists(oldPath) Then ObjFso.DeleteFile oldPath, True
      On Error GoTo 0
      uploadError = "type"
    Else
      Randomize
      stamp = Year(Now) & Right("0" & Month(Now), 2) & Right("0" & Day(Now), 2) & "-" & Right("0" & Hour(Now), 2) & Right("0" & Minute(Now), 2) & Right("0" & Second(Now), 2)
      newName = "case-" & stamp & "-" & CStr(Int((9000 * Rnd) + 1000)) & "." & ext
      newPath = savePath & "\" & newName
      On Error Resume Next
      If ObjFso.FileExists(newPath) Then ObjFso.DeleteFile newPath, True
      ObjFso.MoveFile oldPath, newPath
      If Err.Number <> 0 Then
        Err.Clear
        newName = originalName
      End If
      On Error GoTo 0
      publicPath = "images/case/" & newName
      Call UygulamaAdminRegisterMedia(publicPath, altText, "upload")
    End If
  Else
    targetValue = UygulamaAdminSafeTarget(Request.QueryString("media_target"))
    recordId = 0
    altText = ""
  End If

  Set File1 = Nothing
  Set Upload = Nothing
  Set ObjFso = Nothing

  redirectUrl = "uygulama-ornekleri.asp"
  If recordId > 0 Then redirectUrl = redirectUrl & "?edit=yes&sid=" & recordId & "&" Else redirectUrl = redirectUrl & "?"
  If uploadError = "" Then
    redirectUrl = redirectUrl & "ok=media&media_target=" & Server.URLEncode(targetValue) & "&media_path=" & Server.URLEncode(publicPath) & "&media_alt=" & Server.URLEncode(altText)
  Else
    redirectUrl = redirectUrl & "ok=media_error&err=" & Server.URLEncode(uploadError)
  End If
  Response.Redirect redirectUrl
  Response.End
End Sub

If Request.QueryString("media_upload") = "1" Then
  Call UygulamaAdminHandleMediaUpload()
End If

Call UygulamaAdminEnsureTable()
Call UygulamaAdminEnsureMediaTable()

Dim formAction, recordId, saveRs, deleteId, sortValue, publishValue, slugValue, ctaUrlValue, trackingValue
formAction = Request.Form("form_action")

If formAction = "delete" Then
  deleteId = UygulamaAdminNumber(Request.Form("id"), 0)
  If deleteId > 0 Then baglanti.Execute "DELETE FROM uygulama_ornekleri WHERE [id]=" & deleteId
  Response.Redirect "uygulama-ornekleri.asp?ok=delete"
  Response.End
End If

If formAction = "save" Then
  recordId = UygulamaAdminNumber(Request.Form("id"), 0)
  sortValue = UygulamaAdminNumber(Request.Form("sira"), 0)
  publishValue = 0
  If Request.Form("yayin") = "1" Then publishValue = -1

  slugValue = Trim(Request.Form("slug") & "")
  If slugValue = "" Then slugValue = UygulamaAdminSlug(Request.Form("baslik"))

  trackingValue = Trim(Request.Form("takip_etiketi") & "")
  If trackingValue = "" Then trackingValue = "case_" & Replace(slugValue, "-", "_")

  ctaUrlValue = Trim(Request.Form("cta_url") & "")
  If ctaUrlValue = "" Then ctaUrlValue = "iletisim.asp"

  Set saveRs = Server.CreateObject("ADODB.RecordSet")
  If recordId > 0 Then
    saveRs.Open "SELECT * FROM uygulama_ornekleri WHERE [id]=" & recordId, baglanti, 1, 3
  Else
    saveRs.Open "SELECT * FROM uygulama_ornekleri WHERE [id]=0", baglanti, 1, 3
    saveRs.AddNew
  End If

  If recordId = 0 Or Not saveRs.EOF Then
    Call UygulamaAdminSetField(saveRs, "baslik", Trim(Request.Form("baslik") & ""))
    Call UygulamaAdminSetField(saveRs, "etiket", Trim(Request.Form("etiket") & ""))
    Call UygulamaAdminSetField(saveRs, "kisa_ozet", Trim(Request.Form("kisa_ozet") & ""))
    Call UygulamaAdminSetField(saveRs, "lokasyon", Trim(Request.Form("lokasyon") & ""))
    Call UygulamaAdminSetField(saveRs, "proje_tipi", Trim(Request.Form("proje_tipi") & ""))
    Call UygulamaAdminSetField(saveRs, "hizmet_tipi", Trim(Request.Form("hizmet_tipi") & ""))
    Call UygulamaAdminSetField(saveRs, "sistem_marka", Trim(Request.Form("sistem_marka") & ""))
    Call UygulamaAdminSetField(saveRs, "proje_suresi", Trim(Request.Form("proje_suresi") & ""))
    Call UygulamaAdminSetField(saveRs, "problem", Trim(Request.Form("problem") & ""))
    Call UygulamaAdminSetField(saveRs, "kesif_degerlendirme", Trim(Request.Form("kesif_degerlendirme") & ""))
    Call UygulamaAdminSetField(saveRs, "cozum", Trim(Request.Form("cozum") & ""))
    Call UygulamaAdminSetField(saveRs, "uygulama", Trim(Request.Form("uygulama") & ""))
    Call UygulamaAdminSetField(saveRs, "sonuc", Trim(Request.Form("sonuc") & ""))
    Call UygulamaAdminSetField(saveRs, "cover_image", Trim(Request.Form("cover_image") & ""))
    Call UygulamaAdminSetField(saveRs, "cover_alt", Trim(Request.Form("cover_alt") & ""))
    Call UygulamaAdminSetField(saveRs, "galeri_gorseller", Trim(Request.Form("galeri_gorseller") & ""))
    Call UygulamaAdminSetField(saveRs, "galeri_altlar", Trim(Request.Form("galeri_altlar") & ""))
    Call UygulamaAdminSetField(saveRs, "video_url", Trim(Request.Form("video_url") & ""))
    Call UygulamaAdminSetField(saveRs, "oncesi_image", Trim(Request.Form("oncesi_image") & ""))
    Call UygulamaAdminSetField(saveRs, "oncesi_alt", Trim(Request.Form("oncesi_alt") & ""))
    Call UygulamaAdminSetField(saveRs, "sonrasi_image", Trim(Request.Form("sonrasi_image") & ""))
    Call UygulamaAdminSetField(saveRs, "sonrasi_alt", Trim(Request.Form("sonrasi_alt") & ""))
    Call UygulamaAdminSetField(saveRs, "musteri_yorumu", Trim(Request.Form("musteri_yorumu") & ""))
    Call UygulamaAdminSetField(saveRs, "sss", Trim(Request.Form("sss") & ""))
    Call UygulamaAdminSetField(saveRs, "ilgili_hizmet_url", Trim(Request.Form("ilgili_hizmet_url") & ""))
    Call UygulamaAdminSetField(saveRs, "ilgili_bolge_url", Trim(Request.Form("ilgili_bolge_url") & ""))
    Call UygulamaAdminSetField(saveRs, "cta_metni", Trim(Request.Form("cta_metni") & ""))
    Call UygulamaAdminSetField(saveRs, "cta_url", ctaUrlValue)
    Call UygulamaAdminSetField(saveRs, "slug", slugValue)
    Call UygulamaAdminSetField(saveRs, "takip_etiketi", trackingValue)
    Call UygulamaAdminSetField(saveRs, "seo_title", Trim(Request.Form("seo_title") & ""))
    Call UygulamaAdminSetField(saveRs, "meta_description", Trim(Request.Form("meta_description") & ""))
    Call UygulamaAdminSetField(saveRs, "sira", sortValue)
    Call UygulamaAdminSetField(saveRs, "yayin", publishValue)
    saveRs.Update

    Call UygulamaAdminRegisterMedia(Trim(Request.Form("cover_image") & ""), Trim(Request.Form("cover_alt") & ""), "case")
    Call UygulamaAdminRegisterMedia(Trim(Request.Form("oncesi_image") & ""), Trim(Request.Form("oncesi_alt") & ""), "case")
    Call UygulamaAdminRegisterMedia(Trim(Request.Form("sonrasi_image") & ""), Trim(Request.Form("sonrasi_alt") & ""), "case")
  End If

  If saveRs.State = 1 Then saveRs.Close
  Set saveRs = Nothing
  Response.Redirect "uygulama-ornekleri.asp?ok=save"
  Response.End
End If

Dim editMode, editId, editRs
Dim fBaslik, fEtiket, fKisaOzet, fLokasyon, fProjeTipi, fHizmetTipi, fSistemMarka, fSure
Dim fProblem, fKesif, fCozum, fUygulama, fSonuc
Dim fCoverImage, fCoverAlt, fGaleriGorseller, fGaleriAltlar, fVideoUrl
Dim fOncesiImage, fOncesiAlt, fSonrasiImage, fSonrasiAlt
Dim fMusteriYorumu, fSss, fIlgiliHizmetUrl, fIlgiliBolgeUrl
Dim fCtaMetni, fCtaUrl, fSlug, fTakipEtiketi, fSeoTitle, fMetaDescription, fSira, fYayin
Dim mediaPath, mediaTarget, mediaAlt, mediaNotice

editMode = False
editId = UygulamaAdminNumber(Request.QueryString("sid"), 0)
fBaslik = ""
fEtiket = ""
fKisaOzet = ""
fLokasyon = ""
fProjeTipi = ""
fHizmetTipi = ""
fSistemMarka = ""
fSure = ""
fProblem = ""
fKesif = ""
fCozum = ""
fUygulama = ""
fSonuc = ""
fCoverImage = ""
fCoverAlt = ""
fGaleriGorseller = ""
fGaleriAltlar = ""
fVideoUrl = ""
fOncesiImage = ""
fOncesiAlt = ""
fSonrasiImage = ""
fSonrasiAlt = ""
fMusteriYorumu = ""
fSss = ""
fIlgiliHizmetUrl = ""
fIlgiliBolgeUrl = ""
fCtaMetni = "Benzer bir proje icin kesif talep et"
fCtaUrl = "iletisim.asp"
fSlug = ""
fTakipEtiketi = ""
fSeoTitle = ""
fMetaDescription = ""
fSira = 10
fYayin = True
mediaNotice = ""

If Request.QueryString("edit") = "yes" And editId > 0 Then
  Set editRs = Server.CreateObject("ADODB.RecordSet")
  editRs.Open "SELECT * FROM uygulama_ornekleri WHERE [id]=" & editId, baglanti, 1, 1
  If Not editRs.EOF Then
    editMode = True
    fBaslik = editRs("baslik") & ""
    fEtiket = editRs("etiket") & ""
    fKisaOzet = UygulamaAdminValue(editRs, "kisa_ozet")
    fLokasyon = UygulamaAdminValue(editRs, "lokasyon")
    fProjeTipi = UygulamaAdminValue(editRs, "proje_tipi")
    fHizmetTipi = UygulamaAdminValue(editRs, "hizmet_tipi")
    fSistemMarka = UygulamaAdminValue(editRs, "sistem_marka")
    fSure = UygulamaAdminValue(editRs, "proje_suresi")
    fProblem = UygulamaAdminValue(editRs, "problem")
    fKesif = UygulamaAdminValue(editRs, "kesif_degerlendirme")
    fCozum = UygulamaAdminValue(editRs, "cozum")
    fUygulama = UygulamaAdminValue(editRs, "uygulama")
    fSonuc = UygulamaAdminValue(editRs, "sonuc")
    fCoverImage = UygulamaAdminValue(editRs, "cover_image")
    fCoverAlt = UygulamaAdminValue(editRs, "cover_alt")
    fGaleriGorseller = UygulamaAdminValue(editRs, "galeri_gorseller")
    fGaleriAltlar = UygulamaAdminValue(editRs, "galeri_altlar")
    fVideoUrl = UygulamaAdminValue(editRs, "video_url")
    fOncesiImage = UygulamaAdminValue(editRs, "oncesi_image")
    fOncesiAlt = UygulamaAdminValue(editRs, "oncesi_alt")
    fSonrasiImage = UygulamaAdminValue(editRs, "sonrasi_image")
    fSonrasiAlt = UygulamaAdminValue(editRs, "sonrasi_alt")
    fMusteriYorumu = UygulamaAdminValue(editRs, "musteri_yorumu")
    fSss = UygulamaAdminValue(editRs, "sss")
    fIlgiliHizmetUrl = UygulamaAdminValue(editRs, "ilgili_hizmet_url")
    fIlgiliBolgeUrl = UygulamaAdminValue(editRs, "ilgili_bolge_url")
    fCtaMetni = UygulamaAdminValue(editRs, "cta_metni")
    fCtaUrl = UygulamaAdminValue(editRs, "cta_url")
    fSlug = UygulamaAdminValue(editRs, "slug")
    fTakipEtiketi = UygulamaAdminValue(editRs, "takip_etiketi")
    fSeoTitle = UygulamaAdminValue(editRs, "seo_title")
    fMetaDescription = UygulamaAdminValue(editRs, "meta_description")
    fSira = UygulamaAdminNumber(UygulamaAdminValue(editRs, "sira"), 10)
    fYayin = UygulamaAdminIsLive(UygulamaAdminValue(editRs, "yayin"))
  End If
  If editRs.State = 1 Then editRs.Close
  Set editRs = Nothing
End If

mediaPath = Trim(Request.QueryString("media_path") & "")
mediaTarget = UygulamaAdminSafeTarget(Request.QueryString("media_target"))
mediaAlt = Trim(Request.QueryString("media_alt") & "")
If mediaPath <> "" Then
  Select Case mediaTarget
    Case "cover_image"
      fCoverImage = mediaPath
      If Trim(fCoverAlt) = "" Then fCoverAlt = mediaAlt
    Case "galeri_gorseller"
      If Trim(fGaleriGorseller) = "" Then fGaleriGorseller = mediaPath Else fGaleriGorseller = fGaleriGorseller & vbCrLf & mediaPath
      If mediaAlt <> "" Then
        If Trim(fGaleriAltlar) = "" Then fGaleriAltlar = mediaAlt Else fGaleriAltlar = fGaleriAltlar & vbCrLf & mediaAlt
      End If
    Case "oncesi_image"
      fOncesiImage = mediaPath
      If Trim(fOncesiAlt) = "" Then fOncesiAlt = mediaAlt
    Case "sonrasi_image"
      fSonrasiImage = mediaPath
      If Trim(fSonrasiAlt) = "" Then fSonrasiAlt = mediaAlt
  End Select
  mediaNotice = "Medya forma eklendi. Case kaydina islemek icin Kaydet/Guncelle'ye basin."
End If
%>

<!--#INCLUDE file="menu.asp"-->
<body>

<style>
  .case-admin-wrap, .case-admin-wrap * { box-sizing: border-box; }
  .case-admin-wrap { max-width:1120px; margin:0 auto 50px auto; padding:22px 12px 50px 12px; font-family:Open Sans, Arial, sans-serif; color:#2f2f2f; }
  .case-admin-hero { border:1px solid #dcdcdc; border-radius:18px; background:linear-gradient(135deg,#fff,#f8f1dc); padding:24px; margin-bottom:18px; box-shadow:0 12px 28px rgba(0,0,0,.06); }
  .case-admin-hero h1 { margin:0 0 8px 0; font-size:28px; color:#252525; }
  .case-admin-hero p { margin:0 0 10px 0; color:#666; font-size:14px; line-height:1.6em; }
  .case-admin-schema { display:flex; flex-wrap:wrap; gap:8px; margin-top:12px; }
  .case-admin-schema span { display:inline-flex; border-radius:999px; background:#3c3c3c; color:#f7bb09; padding:6px 10px; font-size:12px; font-weight:800; }
  .case-admin-alert { margin-bottom:16px; padding:12px 14px; border-radius:12px; background:#eef8e8; border:1px solid #cfe7c4; color:#315b20; font-size:13px; font-weight:700; }
  .case-admin-alert.warn { background:#fff4d6; border-color:#eed38a; color:#6b4d00; }
  .case-admin-alert.error { background:#fdecec; border-color:#efb7b7; color:#8a1f1f; }
  .case-admin-card { border:1px solid #dedede; border-radius:16px; background:#fff; margin-bottom:18px; box-shadow:0 10px 24px rgba(0,0,0,.045); overflow:hidden; }
  .case-admin-card-head { display:flex; align-items:center; justify-content:space-between; gap:14px; padding:16px 18px; background:#3c3c3c; color:#fff; }
  .case-admin-card-head strong { color:#f7bb09; font-size:17px; }
  .case-admin-card-head span { color:#f3f3f3; font-size:13px; }
  .case-admin-card-body { padding:18px; }
  .case-admin-grid { display:grid; grid-template-columns:repeat(2,minmax(0,1fr)); gap:14px; }
  .case-admin-grid.three { grid-template-columns:repeat(3,minmax(0,1fr)); }
  .case-admin-field { margin-bottom:14px; }
  .case-admin-field label { display:block; margin-bottom:6px; color:#333; font-size:13px; font-weight:800; }
  .case-admin-field small { display:block; margin-top:5px; color:#777; font-size:12px; line-height:1.45em; }
  .case-admin-field input[type=text], .case-admin-field input[type=file], .case-admin-field select, .case-admin-field textarea { width:100%; border:1px solid #d7d7d7; border-radius:10px; background:#fbfbfb; padding:10px 11px; color:#2f2f2f; font-size:14px; outline:none; }
  .case-admin-field textarea { min-height:92px; resize:vertical; line-height:1.55em; }
  .case-admin-field textarea.tall { min-height:128px; }
  .case-admin-field input:focus, .case-admin-field textarea:focus, .case-admin-field select:focus { border-color:#f7bb09; background:#fff; box-shadow:0 0 0 3px rgba(247,187,9,.18); }
  .case-admin-check { display:inline-flex; align-items:center; gap:8px; padding:10px 12px; border-radius:10px; background:#f7f7f7; font-size:13px; font-weight:800; }
  .case-admin-actions { display:flex; flex-wrap:wrap; align-items:center; gap:10px; margin-top:4px; }
  .case-admin-button, .case-admin-link, .case-admin-danger { display:inline-flex; align-items:center; justify-content:center; border:0; border-radius:10px; padding:10px 16px; font-size:14px; font-weight:800; cursor:pointer; text-decoration:none; }
  .case-admin-button { background:#5e8f2c; color:#fff; }
  .case-admin-link { background:#efefef; color:#333; }
  .case-admin-danger { background:#ba2d2d; color:#fff; }
  .case-admin-table-wrap { overflow-x:auto; }
  .case-admin-table { width:100%; border-collapse:collapse; min-width:980px; font-size:13px; }
  .case-admin-table th { background:#f3f0e4; color:#333; padding:11px 10px; text-align:left; border-bottom:1px solid #dedede; }
  .case-admin-table td { padding:12px 10px; border-bottom:1px solid #ededed; vertical-align:top; }
  .case-admin-status { display:inline-block; border-radius:999px; padding:5px 9px; font-size:12px; font-weight:800; background:#ededed; color:#555; }
  .case-admin-status.live { background:#e7f5de; color:#315b20; }
  .case-media-tools { display:grid; grid-template-columns:330px minmax(0,1fr); gap:16px; align-items:start; }
  .case-media-upload { border:1px solid #ece5d6; border-radius:14px; background:#fffdf7; padding:14px; }
  .case-media-library { display:grid; grid-template-columns:repeat(4,minmax(0,1fr)); gap:10px; max-height:390px; overflow:auto; padding-right:4px; }
  .case-media-item { display:block; width:100%; border:1px solid #e4e4e4; border-radius:12px; background:#fff; padding:0; overflow:hidden; cursor:pointer; text-align:left; }
  .case-media-item:hover { border-color:#f7bb09; box-shadow:0 8px 20px rgba(0,0,0,.08); }
  .case-media-item img { display:block; width:100%; height:88px; object-fit:cover; background:#f2f2f2; }
  .case-media-item span { display:block; padding:8px; color:#555; font-size:11px; line-height:1.3em; word-break:break-all; }
  .case-media-empty { padding:14px; border:1px dashed #d9d9d9; border-radius:12px; color:#666; }
  .case-media-preview { display:grid; grid-template-columns:repeat(4,minmax(0,1fr)); gap:10px; margin-top:8px; }
  .case-media-preview figure { margin:0; border:1px solid #e8e8e8; border-radius:12px; overflow:hidden; background:#fafafa; }
  .case-media-preview img { width:100%; height:92px; object-fit:cover; display:block; }
  .case-media-preview figcaption { padding:7px; color:#666; font-size:11px; word-break:break-all; }
  @media screen and (max-width:900px){ .case-media-tools{grid-template-columns:1fr;} .case-media-library{grid-template-columns:repeat(2,minmax(0,1fr));} }
  @media screen and (max-width:760px){ .case-admin-grid,.case-admin-grid.three{grid-template-columns:1fr;} .case-admin-card-head{display:block;} .case-media-preview{grid-template-columns:1fr 1fr;} }
</style>

<div class="baslik">Uygulama Ornekleri</div>

<div class="case-admin-wrap">
  <div class="case-admin-hero">
    <h1>Case study paneli</h1>
    <p>Bu panel kart + detay sayfali mini case sistemi icin calisir. Medya alanlari artik cihazdan yukleme ve galeriden secme akisiyle doldurulabilir.</p>
    <p>Guncel alan yapisi: baslik, etiket, kisa ozet, lokasyon, proje tipi, hizmet tipi, sistem/marka, sure, problem, kesif/degerlendirme, cozum, uygulama, sonuc, CTA, slug, SEO, kapak, galeri, video, oncesi/sonrasi, musteri notu ve SSS.</p>
    <div class="case-admin-schema">
      <span>Cihazdan yukle</span><span>Galeriden sec</span><span>DB medya kutuphanesi</span><span>Kapak foto</span><span>Galeri</span><span>Oncesi / sonrasi</span><span>Alt text</span>
    </div>
  </div>

  <% If Request.QueryString("ok") = "save" Then %>
    <div class="case-admin-alert">Kayit basariyla kaydedildi.</div>
  <% ElseIf Request.QueryString("ok") = "delete" Then %>
    <div class="case-admin-alert">Kayit silindi.</div>
  <% ElseIf Request.QueryString("ok") = "media" Or mediaNotice <> "" Then %>
    <div class="case-admin-alert warn"><%=UygulamaAdminHtml(mediaNotice)%></div>
  <% ElseIf Request.QueryString("ok") = "media_error" Then %>
    <div class="case-admin-alert error">Medya yuklenemedi. Hata: <%=UygulamaAdminHtml(Request.QueryString("err"))%></div>
  <% End If %>

  <div class="case-admin-card">
    <div class="case-admin-card-head">
      <strong>Medya yukle / galeriden sec</strong>
      <span>Secilen yol asagidaki form alanina yazilir; case'e islemek icin kaydetmek gerekir.</span>
    </div>
    <div class="case-admin-card-body">
      <div class="case-media-tools">
        <form method="post" action="uygulama-ornekleri.asp?media_upload=1" enctype="multipart/form-data" class="case-media-upload">
          <input type="hidden" name="id" value="<%=editId%>">
          <div class="case-admin-field">
            <label>Yuklenecek alan</label>
            <select name="media_target">
              <option value="cover_image">Kapak foto</option>
              <option value="galeri_gorseller">Galeriye ekle</option>
              <option value="oncesi_image">Oncesi foto</option>
              <option value="sonrasi_image">Sonrasi foto</option>
            </select>
          </div>
          <div class="case-admin-field">
            <label>Cihazdan gorsel sec</label>
            <input type="file" name="resim" accept="image/*" required>
            <small>JPG, PNG, GIF veya WEBP. Dosya images/case/ klasorune kaydedilir.</small>
          </div>
          <div class="case-admin-field">
            <label>Alt text / not</label>
            <input type="text" name="alt_text" maxlength="255" placeholder="Lara villa klima montaji kapak fotografi">
          </div>
          <button type="submit" class="case-admin-button">Yukle ve forma ekle</button>
        </form>

        <div>
          <div class="case-admin-grid">
            <div class="case-admin-field">
              <label>Galeriden secilince nereye eklensin?</label>
              <select id="caseMediaTarget">
                <option value="cover_image">Kapak foto</option>
                <option value="galeri_gorseller">Galeriye ekle</option>
                <option value="oncesi_image">Oncesi foto</option>
                <option value="sonrasi_image">Sonrasi foto</option>
              </select>
              <small>Asagidaki kutuphaneden gorsele tiklayinca bu alana yazilir.</small>
            </div>
          </div>
          <% Call UygulamaAdminRenderMediaLibrary() %>
        </div>
      </div>
    </div>
  </div>

  <div class="case-admin-card">
    <div class="case-admin-card-head">
      <strong><% If editMode Then %>Case'i duzenle<% Else %>Yeni case ekle<% End If %></strong>
      <span>Problem &rarr; Kesif &rarr; Cozum &rarr; Uygulama &rarr; Sonuc &rarr; CTA akisi</span>
    </div>
    <div class="case-admin-card-body">
      <form method="post" action="uygulama-ornekleri.asp">
        <input type="hidden" name="form_action" value="save">
        <input type="hidden" name="id" value="<%=editId%>">

        <div class="case-admin-grid">
          <div class="case-admin-field"><label>Baslik</label><input type="text" name="baslik" value="<%=UygulamaAdminHtml(fBaslik)%>" maxlength="255" required></div>
          <div class="case-admin-field"><label>Etiket</label><input type="text" name="etiket" value="<%=UygulamaAdminHtml(fEtiket)%>" maxlength="100" placeholder="Villa / VRF / Bakim"></div>
        </div>

        <div class="case-admin-field"><label>Kisa ozet</label><textarea name="kisa_ozet"><%=UygulamaAdminTextarea(fKisaOzet)%></textarea><small>Kartta ve detay sayfasi girisinde gorunur. 1-2 cumle yeterli.</small></div>

        <div class="case-admin-grid three">
          <div class="case-admin-field"><label>Lokasyon</label><input type="text" name="lokasyon" value="<%=UygulamaAdminHtml(fLokasyon)%>" maxlength="150" placeholder="Lara / Muratpasa / Antalya"></div>
          <div class="case-admin-field"><label>Proje tipi</label><input type="text" name="proje_tipi" value="<%=UygulamaAdminHtml(fProjeTipi)%>" maxlength="120" placeholder="Otel / Ofis / Villa"></div>
          <div class="case-admin-field"><label>Hizmet tipi</label><input type="text" name="hizmet_tipi" value="<%=UygulamaAdminHtml(fHizmetTipi)%>" maxlength="150" placeholder="VRF sistemleri / Klima montaji"></div>
        </div>

        <div class="case-admin-grid">
          <div class="case-admin-field"><label>Sistem / marka</label><input type="text" name="sistem_marka" value="<%=UygulamaAdminHtml(fSistemMarka)%>" maxlength="180" placeholder="Mitsubishi Electric VRF"></div>
          <div class="case-admin-field"><label>Proje suresi</label><input type="text" name="proje_suresi" value="<%=UygulamaAdminHtml(fSure)%>" maxlength="80" placeholder="1 gun / 3 gun / Ayni gun"></div>
        </div>

        <div class="case-admin-grid">
          <div class="case-admin-field"><label>Problem</label><textarea name="problem" class="tall"><%=UygulamaAdminTextarea(fProblem)%></textarea></div>
          <div class="case-admin-field"><label>Kesif / degerlendirme</label><textarea name="kesif_degerlendirme" class="tall"><%=UygulamaAdminTextarea(fKesif)%></textarea><small>m2, cephe, drenaj, dis unite, elektrik ve servis erisimi gibi teknik dusunce burada.</small></div>
        </div>

        <div class="case-admin-grid">
          <div class="case-admin-field"><label>Cozum</label><textarea name="cozum" class="tall"><%=UygulamaAdminTextarea(fCozum)%></textarea></div>
          <div class="case-admin-field"><label>Uygulama</label><textarea name="uygulama" class="tall"><%=UygulamaAdminTextarea(fUygulama)%></textarea></div>
        </div>

        <div class="case-admin-grid">
          <div class="case-admin-field"><label>Sonuc</label><textarea name="sonuc" class="tall"><%=UygulamaAdminTextarea(fSonuc)%></textarea></div>
          <div class="case-admin-field"><label>Musteri yorumu / notu</label><textarea name="musteri_yorumu" class="tall"><%=UygulamaAdminTextarea(fMusteriYorumu)%></textarea><small>Opsiyonel. Isim vermeden anonim not da girilebilir.</small></div>
        </div>

        <div class="case-admin-card" style="box-shadow:none; margin-top:10px">
          <div class="case-admin-card-head"><strong>Medya alanlari</strong><span>Yukle/sec paneli bu alanlari doldurur; istersen elle de duzenleyebilirsin.</span></div>
          <div class="case-admin-card-body">
            <div class="case-admin-grid">
              <div class="case-admin-field"><label>Kapak foto URL / dosya yolu</label><input type="text" id="cover_image" name="cover_image" value="<%=UygulamaAdminHtml(fCoverImage)%>" maxlength="255" placeholder="images/case/lara-villa.webp"></div>
              <div class="case-admin-field"><label>Kapak alt text</label><input type="text" id="cover_alt" name="cover_alt" value="<%=UygulamaAdminHtml(fCoverAlt)%>" maxlength="255" placeholder="Lara villa klima montaji"></div>
            </div>
            <div class="case-admin-grid">
              <div class="case-admin-field"><label>Galeri foto URLleri</label><textarea id="galeri_gorseller" name="galeri_gorseller"><%=UygulamaAdminTextarea(fGaleriGorseller)%></textarea><small>Her satira bir foto yolu.</small></div>
              <div class="case-admin-field"><label>Galeri alt textleri</label><textarea id="galeri_altlar" name="galeri_altlar"><%=UygulamaAdminTextarea(fGaleriAltlar)%></textarea><small>Her satir, soldaki ayni siradaki fotografin alt metni.</small></div>
            </div>
            <div class="case-admin-field"><label>Video URL</label><input type="text" name="video_url" value="<%=UygulamaAdminHtml(fVideoUrl)%>" maxlength="255" placeholder="YouTube / Vimeo / Instagram linki - opsiyonel"></div>
            <div class="case-admin-grid">
              <div class="case-admin-field"><label>Oncesi foto</label><input type="text" id="oncesi_image" name="oncesi_image" value="<%=UygulamaAdminHtml(fOncesiImage)%>" maxlength="255" placeholder="Opsiyonel"></div>
              <div class="case-admin-field"><label>Oncesi alt text</label><input type="text" id="oncesi_alt" name="oncesi_alt" value="<%=UygulamaAdminHtml(fOncesiAlt)%>" maxlength="255"></div>
            </div>
            <div class="case-admin-grid">
              <div class="case-admin-field"><label>Sonrasi foto</label><input type="text" id="sonrasi_image" name="sonrasi_image" value="<%=UygulamaAdminHtml(fSonrasiImage)%>" maxlength="255" placeholder="Opsiyonel"></div>
              <div class="case-admin-field"><label>Sonrasi alt text</label><input type="text" id="sonrasi_alt" name="sonrasi_alt" value="<%=UygulamaAdminHtml(fSonrasiAlt)%>" maxlength="255"></div>
            </div>
            <div class="case-media-preview" id="caseMediaPreview"></div>
          </div>
        </div>

        <div class="case-admin-grid">
          <div class="case-admin-field"><label>SSS</label><textarea name="sss" class="tall"><%=UygulamaAdminTextarea(fSss)%></textarea><small>Format: Soru|Cevap. Her satira bir SSS.</small></div>
          <div class="case-admin-field"><label>SEO meta description</label><textarea name="meta_description" class="tall"><%=UygulamaAdminTextarea(fMetaDescription)%></textarea><small>Bos kalirsa kisa ozet/problem kullanilir.</small></div>
        </div>

        <div class="case-admin-grid">
          <div class="case-admin-field"><label>CTA metni</label><input type="text" name="cta_metni" value="<%=UygulamaAdminHtml(fCtaMetni)%>" maxlength="255"></div>
          <div class="case-admin-field"><label>CTA linki</label><input type="text" name="cta_url" value="<%=UygulamaAdminHtml(fCtaUrl)%>" maxlength="255" placeholder="iletisim.asp"><small>Takip icin /iletisim.asp?source=case_x kullanilabilir.</small></div>
        </div>

        <div class="case-admin-grid">
          <div class="case-admin-field"><label>Slug</label><input type="text" name="slug" value="<%=UygulamaAdminHtml(fSlug)%>" maxlength="120" placeholder="Bos kalirsa basliktan uretilir"><small>Detay URL: ../uygulama-ornegi.asp?slug=<%=UygulamaAdminHtml(fSlug)%></small></div>
          <div class="case-admin-field"><label>Takip etiketi</label><input type="text" name="takip_etiketi" value="<%=UygulamaAdminHtml(fTakipEtiketi)%>" maxlength="120" placeholder="case_lara_otel_vrf"></div>
        </div>

        <div class="case-admin-grid">
          <div class="case-admin-field"><label>SEO title</label><input type="text" name="seo_title" value="<%=UygulamaAdminHtml(fSeoTitle)%>" maxlength="255" placeholder="Bos kalirsa baslik + Ozum Klima kullanilir"></div>
          <div class="case-admin-field"><label>Siralama</label><input type="text" name="sira" value="<%=UygulamaAdminHtml(fSira)%>" maxlength="6"></div>
        </div>

        <div class="case-admin-grid">
          <div class="case-admin-field"><label>Ilgili hizmet URL</label><input type="text" name="ilgili_hizmet_url" value="<%=UygulamaAdminHtml(fIlgiliHizmetUrl)%>" maxlength="255" placeholder="antalya-vrf-sistemleri.asp"></div>
          <div class="case-admin-field"><label>Ilgili bolge URL</label><input type="text" name="ilgili_bolge_url" value="<%=UygulamaAdminHtml(fIlgiliBolgeUrl)%>" maxlength="255" placeholder="muratpasa-klima-servisi.asp"></div>
        </div>

        <div class="case-admin-field"><label class="case-admin-check"><input type="checkbox" name="yayin" value="1" <% If fYayin Then Response.Write "checked" %>> Yayinda goster</label></div>

        <div class="case-admin-actions">
          <button type="submit" class="case-admin-button"><% If editMode Then %>Guncelle<% Else %>Kaydet<% End If %></button>
          <% If editMode Then %><a href="uygulama-ornekleri.asp" class="case-admin-link">Yeni kayit moduna don</a><% End If %>
          <a href="../uygulama-ornekleri.asp" target="_blank" class="case-admin-link">Listeyi sitede gor</a>
          <% If editMode And Trim(fSlug) <> "" Then %><a href="../uygulama-ornegi.asp?slug=<%=Server.URLEncode(fSlug)%>" target="_blank" class="case-admin-link">Detay sayfasini ac</a><% End If %>
        </div>
      </form>
    </div>
  </div>

  <div class="case-admin-card">
    <div class="case-admin-card-head"><strong>Kayit listesi</strong><span>Sira numarasina gore yayinda gorunur; detay linki slug ile olusur.</span></div>
    <div class="case-admin-card-body case-admin-table-wrap">
      <%
      Dim listRs, listSlug, hasMedia
      Set listRs = Server.CreateObject("ADODB.RecordSet")
      listRs.Open "SELECT * FROM uygulama_ornekleri ORDER BY [sira] ASC, [id] ASC", baglanti, 1, 1
      %>
      <table class="case-admin-table">
        <tr><th width="55">Sira</th><th>Baslik / slug</th><th width="150">Lokasyon</th><th width="140">Hizmet</th><th width="95">Medya</th><th width="110">Durum</th><th width="90">Duzenle</th><th width="80">Detay</th><th width="80">Sil</th></tr>
        <% If listRs.EOF Then %>
          <tr><td colspan="9">Henuz kayit yok.</td></tr>
        <% Else %>
          <% Do While Not listRs.EOF
            listSlug = UygulamaAdminValue(listRs, "slug")
            hasMedia = "Yok"
            If Trim(UygulamaAdminValue(listRs, "cover_image")) <> "" Or Trim(UygulamaAdminValue(listRs, "galeri_gorseller")) <> "" Or Trim(UygulamaAdminValue(listRs, "video_url")) <> "" Then hasMedia = "Var"
          %>
            <tr>
              <td><%=UygulamaAdminHtml(UygulamaAdminValue(listRs, "sira"))%></td>
              <td><b><%=UygulamaAdminHtml(listRs("baslik"))%></b><br><span style="color:#777"><%=UygulamaAdminHtml(listSlug)%></span></td>
              <td><%=UygulamaAdminHtml(UygulamaAdminValue(listRs, "lokasyon"))%></td>
              <td><%=UygulamaAdminHtml(UygulamaAdminValue(listRs, "hizmet_tipi"))%></td>
              <td><%=hasMedia%></td>
              <td><% If UygulamaAdminIsLive(UygulamaAdminValue(listRs, "yayin")) Then %><span class="case-admin-status live">Yayinda</span><% Else %><span class="case-admin-status">Pasif</span><% End If %></td>
              <td><a class="case-admin-link" href="uygulama-ornekleri.asp?edit=yes&sid=<%=listRs("id")%>">Duzenle</a></td>
              <td><% If Trim(listSlug) <> "" Then %><a class="case-admin-link" target="_blank" href="../uygulama-ornegi.asp?slug=<%=Server.URLEncode(listSlug)%>">Ac</a><% End If %></td>
              <td><form method="post" action="uygulama-ornekleri.asp" style="margin:0"><input type="hidden" name="form_action" value="delete"><input type="hidden" name="id" value="<%=listRs("id")%>"><button type="submit" class="case-admin-danger" onclick="return confirm('Bu uygulama ornegini silmek istiyor musunuz?')">Sil</button></form></td>
            </tr>
          <% listRs.MoveNext : Loop %>
        <% End If %>
      </table>
      <% If listRs.State = 1 Then listRs.Close : Set listRs = Nothing %>
    </div>
  </div>
</div>

<script>
function caseMediaToPreviewSrc(path) {
  if (!path) return '';
  if (/^https?:\/\//i.test(path) || path.indexOf('/') === 0 || path.indexOf('../') === 0) return path;
  return '../' + path;
}
function casePickMedia(path) {
  var target = document.getElementById('caseMediaTarget').value;
  var input = document.getElementById(target);
  if (!input || !path) return;
  if (target === 'galeri_gorseller') {
    input.value = input.value ? input.value.replace(/\s+$/,'') + '\n' + path : path;
  } else {
    input.value = path;
  }
  caseRefreshMediaPreview();
}
function caseRefreshMediaPreview() {
  var wrap = document.getElementById('caseMediaPreview');
  if (!wrap) return;
  var items = [];
  ['cover_image','oncesi_image','sonrasi_image'].forEach(function(id){
    var el = document.getElementById(id);
    if (el && el.value) items.push({label:id, path:el.value});
  });
  var gallery = document.getElementById('galeri_gorseller');
  if (gallery && gallery.value) {
    gallery.value.split(/\r?\n/).forEach(function(line, index){
      line = line.replace(/^\s+|\s+$/g,'');
      if (line) items.push({label:'galeri ' + (index + 1), path:line});
    });
  }
  wrap.innerHTML = '';
  items.slice(0,8).forEach(function(item){
    var fig = document.createElement('figure');
    var img = document.createElement('img');
    var cap = document.createElement('figcaption');
    img.src = caseMediaToPreviewSrc(item.path);
    cap.appendChild(document.createTextNode(item.label + ': ' + item.path));
    fig.appendChild(img);
    fig.appendChild(cap);
    wrap.appendChild(fig);
  });
}
['cover_image','galeri_gorseller','oncesi_image','sonrasi_image'].forEach(function(id){
  var el = document.getElementById(id);
  if (el) el.oninput = caseRefreshMediaPreview;
});
caseRefreshMediaPreview();
</script>

</body>