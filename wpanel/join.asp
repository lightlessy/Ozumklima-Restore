<%@ Language="VBScript" CodePage="65001" %>
<%
Response.Buffer = True
Response.CodePage = 65001
Response.Charset = "utf-8"
%>

<!--#include file="menu.asp"-->

<body onload="initialiseWebWizRTE();">

<div class="baslik">Ürün Ekle</div>

<script src="../js/jscolor.js"></script>

<%
Dim grup, tip, surface, actionName, formMessage, formMessageType
Dim Rs, mods, Sorgula, yeniKayitID

grup = Trim(Request.QueryString("g") & "")
tip = Trim(Request.QueryString("tip") & "")
surface = Trim(Request.QueryString("surface") & "")
actionName = LCase(Trim(Request.QueryString("action") & ""))
formMessage = ""
formMessageType = ""

If grup <> "" And Not IsNumeric(grup) Then grup = ""
If tip <> "" And Not IsNumeric(tip) Then tip = ""
If surface <> "" And Not IsNumeric(surface) Then surface = ""

If grup <> "" Then grup = CStr(CLng(grup))
If tip <> "" Then tip = CStr(CLng(tip))
If surface <> "" Then surface = CStr(CLng(surface))

If UCase(Request.ServerVariables("REQUEST_METHOD")) = "POST" Then

    If actionName = "add-main-category" Then
        Dim anaKategoriAdi
        anaKategoriAdi = Trim(Request.Form("isim1") & "")

        If anaKategoriAdi = "" Then
            formMessage = "Ana kategori adı boş bırakılamaz."
            formMessageType = "error"
        Else
            On Error Resume Next

            Set Rs = Server.CreateObject("ADODB.Recordset")
            Rs.Open "SELECT * FROM anagrup", baglanti, 1, 3
            Rs.AddNew
            Rs("isim") = anaKategoriAdi
            Rs("durum") = 1
            Rs.Update

            If Err.Number <> 0 Then
                formMessage = "Ana kategori eklenemedi: " & Err.Description
                formMessageType = "error"
                Err.Clear
            Else
                yeniKayitID = Rs("id")
            End If

            If Rs.State = 1 Then Rs.Close
            Set Rs = Nothing
            On Error GoTo 0

            If formMessageType <> "error" Then
                Response.Redirect "join.asp?g=" & CLng(yeniKayitID) & "&msg=main-added"
                Response.End
            End If
        End If
    End If

    If actionName = "add-type" Then
        Dim tipAdi, tipGrup
        tipAdi = Trim(Request.Form("isim2") & "")
        tipGrup = Trim(Request.Form("grup") & "")

        If tipAdi = "" Then
            formMessage = "Alt kategori adı boş bırakılamaz."
            formMessageType = "error"
        ElseIf Not IsNumeric(tipGrup) Then
            formMessage = "Alt kategori için geçerli bir ana kategori seçilmelidir."
            formMessageType = "error"
        Else
            On Error Resume Next

            Set Rs = Server.CreateObject("ADODB.Recordset")
            Rs.Open "SELECT * FROM tip", baglanti, 1, 3
            Rs.AddNew
            Rs("isim") = tipAdi
            Rs("grup") = CLng(tipGrup)
            Rs.Update

            If Err.Number <> 0 Then
                formMessage = "Alt kategori eklenemedi: " & Err.Description
                formMessageType = "error"
                Err.Clear
            Else
                yeniKayitID = Rs("id")
            End If

            If Rs.State = 1 Then Rs.Close
            Set Rs = Nothing
            On Error GoTo 0

            If formMessageType <> "error" Then
                Response.Redirect "join.asp?g=" & CLng(tipGrup) & "&tip=" & CLng(yeniKayitID) & "&msg=type-added"
                Response.End
            End If
        End If
    End If

    If actionName = "add-surface" Then
        Dim surfaceAdi, surfaceGrup, surfaceTip
        surfaceAdi = Trim(Request.Form("isim3") & "")
        surfaceGrup = Trim(Request.Form("grubu") & "")
        surfaceTip = Trim(Request.Form("selectedTip") & "")

        If surfaceAdi = "" Then
            formMessage = "Ürün grubu adı boş bırakılamaz."
            formMessageType = "error"
        ElseIf Not IsNumeric(surfaceGrup) Then
            formMessage = "Ürün grubu için geçerli bir ana kategori seçilmelidir."
            formMessageType = "error"
        ElseIf Not IsNumeric(surfaceTip) Then
            formMessage = "Ürün grubu için geçerli bir alt kategori seçilmelidir."
            formMessageType = "error"
        Else
            On Error Resume Next

            Set Rs = Server.CreateObject("ADODB.Recordset")
            Rs.Open "SELECT * FROM surface", baglanti, 1, 3
            Rs.AddNew
            Rs("isim") = surfaceAdi
            Rs("grup") = CLng(surfaceGrup)
            Rs.Update

            If Err.Number <> 0 Then
                formMessage = "Ürün grubu eklenemedi: " & Err.Description
                formMessageType = "error"
                Err.Clear
            Else
                yeniKayitID = Rs("id")
            End If

            If Rs.State = 1 Then Rs.Close
            Set Rs = Nothing
            On Error GoTo 0

            If formMessageType <> "error" Then
                Response.Redirect "join.asp?g=" & CLng(surfaceGrup) & "&tip=" & CLng(surfaceTip) & "&surface=" & CLng(yeniKayitID) & "&msg=surface-added"
                Response.End
            End If
        End If
    End If

End If

Select Case LCase(Trim(Request.QueryString("msg") & ""))
    Case "main-added"
        formMessage = "Ana kategori eklendi. Şimdi alt kategoriyi seçebilir veya ekleyebilirsiniz."
        formMessageType = "success"
    Case "type-added"
        formMessage = "Alt kategori eklendi. Ürün bilgilerini doldurabilirsiniz."
        formMessageType = "success"
    Case "surface-added"
        formMessage = "Ürün grubu eklendi ve seçildi."
        formMessageType = "success"
    Case "product-error"
        formMessage = "Ürün kaydedilemedi. Alanları kontrol edip tekrar deneyin."
        formMessageType = "error"
End Select
%>

<style>
.product-add-wrap {
    width: 94%;
    max-width: 980px;
    margin: 28px auto 60px;
    font-family: Arial, Helvetica, sans-serif;
    font-size: 14px;
    color: #353535;
}
.product-add-card {
    background: #fff;
    border: 1px solid #dedede;
    box-shadow: 0 2px 8px rgba(0,0,0,.06);
    margin-bottom: 18px;
}
.product-add-card__head {
    padding: 15px 18px;
    border-bottom: 1px solid #e8e8e8;
    background: #f7f4f0;
}
.product-add-card__head strong {
    font-size: 16px;
}
.product-add-card__body {
    padding: 18px;
}
.product-add-grid {
    display: grid;
    grid-template-columns: 180px minmax(0, 1fr);
    gap: 13px 18px;
    align-items: center;
}
.product-add-grid label {
    font-weight: bold;
}
.product-add-grid input[type="text"],
.product-add-grid input[type="number"],
.product-add-grid select,
.product-add-grid textarea {
    width: 100%;
    box-sizing: border-box;
    border: 1px solid #cfcfcf;
    background: #fff;
    padding: 9px 10px;
    font: inherit;
}
.product-add-grid textarea {
    resize: vertical;
}
.product-add-inline {
    display: flex;
    gap: 8px;
    align-items: center;
}
.product-add-inline > :first-child {
    flex: 1;
}
.product-add-button {
    display: inline-block;
    border: 1px solid #333;
    background: #3c3c3c;
    color: #fff;
    padding: 9px 14px;
    cursor: pointer;
    font-weight: bold;
    white-space: nowrap;
}
.product-add-button:hover {
    background: #555;
}
.product-add-button--primary {
    padding: 12px 26px;
    font-size: 15px;
}
.product-add-button--secondary {
    border-color: #b88a28;
    background: #b88a28;
}
.product-add-help {
    color: #747474;
    font-size: 12px;
    margin-top: 5px;
    line-height: 1.45;
}
.product-add-message {
    padding: 12px 15px;
    margin-bottom: 18px;
    border: 1px solid #d8d8d8;
    background: #f7f7f7;
}
.product-add-message--success {
    border-color: #9fc7a6;
    background: #f1faf2;
    color: #276533;
}
.product-add-message--error {
    border-color: #d7a0a0;
    background: #fff3f3;
    color: #8b2424;
}
.product-add-step {
    display: inline-block;
    min-width: 22px;
    height: 22px;
    line-height: 22px;
    margin-right: 7px;
    border-radius: 50%;
    text-align: center;
    color: #fff;
    background: #3c3c3c;
    font-size: 12px;
}
.product-add-colors {
    display: flex;
    flex-wrap: wrap;
    gap: 8px;
}
.product-add-colors input {
    width: 76px !important;
}
.product-add-actions {
    text-align: center;
    padding-top: 5px;
}
.product-add-empty {
    padding: 24px;
    border: 1px dashed #c9c9c9;
    background: #fafafa;
    text-align: center;
    color: #666;
}
.product-add-check {
    display: flex;
    align-items: center;
    gap: 8px;
}
.product-add-check label {
    font-weight: normal;
}
@media (max-width: 760px) {
    .product-add-wrap {
        width: 96%;
    }
    .product-add-grid {
        grid-template-columns: 1fr;
        gap: 7px;
    }
    .product-add-inline {
        align-items: stretch;
        flex-direction: column;
    }
    .product-add-button {
        width: 100%;
        box-sizing: border-box;
    }
}
</style>

<div class="product-add-wrap">

<% If formMessage <> "" Then %>
    <div class="product-add-message product-add-message--<%=Server.HTMLEncode(formMessageType)%>">
        <%=Server.HTMLEncode(formMessage)%>
    </div>
<% End If %>

<div class="product-add-card">
    <div class="product-add-card__head">
        <strong><span class="product-add-step">1</span>Kategori seçimi</strong>
    </div>
    <div class="product-add-card__body">
        <div class="product-add-grid">
            <label for="anagrup">Ürün kategorisi</label>
            <div>
                <select id="anagrup" name="anagrupPicker">
                    <option value="">Lütfen seçiniz</option>
<%
Set mods = Server.CreateObject("ADODB.Recordset")
mods.Open "SELECT id, isim FROM anagrup ORDER BY sira DESC, isim ASC", baglanti, 1, 1
Do While Not mods.EOF
%>
                    <option value="<%=CLng(mods("id"))%>"<% If grup <> "" Then If CLng(grup) = CLng(mods("id")) Then Response.Write " selected" %>><%=Server.HTMLEncode(mods("isim") & "")%></option>
<%
    mods.MoveNext
Loop
mods.Close
Set mods = Nothing
%>
                </select>
                <div class="product-add-help">Kategori değiştiğinde alt kategori listesi otomatik yenilenir.</div>
            </div>

            <label for="isim1">Yeni ana kategori</label>
            <form method="post" action="join.asp?action=add-main-category" class="product-add-inline">
                <input type="text" id="isim1" name="isim1" maxlength="150" placeholder="Kategori adı">
                <button type="submit" class="product-add-button product-add-button--secondary">+ Ana kategori ekle</button>
            </form>

<% If grup <> "" Then %>
            <label for="tip">Alt kategori</label>
            <div>
                <select id="tip" name="tipPicker">
                    <option value="">Lütfen seçiniz</option>
<%
Set mods = Server.CreateObject("ADODB.Recordset")
mods.Open "SELECT id, isim FROM tip WHERE grup=" & CLng(grup) & " ORDER BY sira DESC, isim ASC", baglanti, 1, 1
Do While Not mods.EOF
%>
                    <option value="<%=CLng(mods("id"))%>"<% If tip <> "" Then If CLng(tip) = CLng(mods("id")) Then Response.Write " selected" %>><%=Server.HTMLEncode(mods("isim") & "")%></option>
<%
    mods.MoveNext
Loop
mods.Close
Set mods = Nothing
%>
                </select>
            </div>

            <label for="isim2">Yeni alt kategori</label>
            <form method="post" action="join.asp?g=<%=CLng(grup)%>&action=add-type" class="product-add-inline">
                <input type="hidden" name="grup" value="<%=CLng(grup)%>">
                <input type="text" id="isim2" name="isim2" maxlength="150" placeholder="Alt kategori adı">
                <button type="submit" class="product-add-button product-add-button--secondary">+ Alt kategori ekle</button>
            </form>
<% End If %>
        </div>
    </div>
</div>

<% If grup <> "" And tip <> "" Then %>
<form id="productForm" method="post" action="join2.asp">
    <input type="hidden" name="anagrup" value="<%=CLng(grup)%>">
    <input type="hidden" name="tip" value="<%=CLng(tip)%>">

    <div class="product-add-card">
        <div class="product-add-card__head">
            <strong><span class="product-add-step">2</span>Temel ürün bilgileri</strong>
        </div>
        <div class="product-add-card__body">
            <div class="product-add-grid">
                <label for="surface">Ürün grubu</label>
                <div class="product-add-inline">
                    <select name="surface" id="surface">
                        <option value="0">Yok / seçilmedi</option>
<%
Set mods = Server.CreateObject("ADODB.Recordset")
mods.Open "SELECT id, isim FROM surface WHERE grup=" & CLng(grup) & " ORDER BY isim ASC", baglanti, 1, 1
Do While Not mods.EOF
%>
                        <option value="<%=CLng(mods("id"))%>"<% If surface <> "" Then If CLng(surface) = CLng(mods("id")) Then Response.Write " selected" %>><%=Server.HTMLEncode(mods("isim") & "")%></option>
<%
    mods.MoveNext
Loop
mods.Close
Set mods = Nothing
%>
                    </select>
                </div>

                <label for="isim3">Yeni ürün grubu</label>
                <div class="product-add-inline">
                    <input type="text" id="isim3" name="surfaceDraft" maxlength="150" placeholder="Ürün grubu adı">
                    <button type="button" id="addSurfaceButton" class="product-add-button product-add-button--secondary">+ Ürün grubu ekle</button>
                </div>

                <label for="kodu">Ürün kodu</label>
                <input type="text" id="kodu" name="kodu" maxlength="100" value="0">

                <label for="isim">Ürün adı</label>
                <input type="text" id="isim" name="isim" maxlength="255" required placeholder="Örn. Mitsubishi Electric MSZ-AY35VGK">

                <label for="descr">Kısa açıklama</label>
                <div>
                    <textarea id="descr" name="descr" rows="4" maxlength="2000" placeholder="Listeleme ve arama sonuçlarında kullanılacak kısa ürün açıklaması"></textarea>
                    <div class="product-add-help">Ürünün ana özelliklerini 1-3 cümlede özetleyin.</div>
                </div>

                <label for="delivery">Teslim / kargo süresi</label>
                <div>
                    <input type="number" id="delivery" name="delivery" min="0" max="365" value="3">
                    <div class="product-add-help">Gün olarak girin.</div>
                </div>

                <label>Yayın durumu</label>
                <div class="product-add-check">
                    <input type="checkbox" id="yayin" name="yayin" value="1">
                    <label for="yayin">Ürünü kaydeder kaydetmez yayına al</label>
                </div>
            </div>
        </div>
    </div>

    <div class="product-add-card">
        <div class="product-add-card__head">
            <strong><span class="product-add-step">3</span>İçerik ve arama bilgileri</strong>
        </div>
        <div class="product-add-card__body">
            <div class="product-add-grid">
                <label for="notlar">Ürün detayları</label>
                <div>
<%
strFormName = "productForm"
strTextAreaName = "notlar"
%>
                    <!--#include file="RTE_editor_inc.asp"-->
                    <textarea id="notlar" name="notlar" rows="18" style="margin-top:8px;"></textarea>
                </div>

                <label for="keyw">Anahtar kelimeler</label>
                <div>
                    <textarea id="keyw" name="keyw" rows="4" maxlength="2000" placeholder="klima, inverter klima, Mitsubishi Electric..."></textarea>
                    <div class="product-add-help">Arama sayfası ürün adı, kısa açıklama ve bu alan içinde eşleşme arar.</div>
                </div>

                <label>Renk seçenekleri</label>
                <div>
                    <div class="product-add-colors">
<%
Dim renkIndex
For renkIndex = 1 To 5
%>
                        <input class="jscolor" name="renk<%=renkIndex%>" value="0" aria-label="Renk <%=renkIndex%>">
<%
Next
%>
                    </div>
                    <div class="product-add-help">Kullanılmayan renkleri 0 olarak bırakabilirsiniz.</div>
                </div>
            </div>
        </div>
    </div>

    <div class="product-add-actions">
        <button type="submit" id="productSubmit" class="product-add-button product-add-button--primary">Ürünü kaydet ve görsel ekle</button>
    </div>
</form>

<form id="surfaceForm" method="post" action="join.asp?g=<%=CLng(grup)%>&tip=<%=CLng(tip)%>&action=add-surface" style="display:none;">
    <input type="hidden" name="grubu" value="<%=CLng(grup)%>">
    <input type="hidden" name="selectedTip" value="<%=CLng(tip)%>">
    <input type="hidden" name="isim3" id="surfaceNameHidden" value="">
</form>

<% Else %>
    <div class="product-add-empty">
        Ürün formunu açmak için önce ana kategori ve alt kategori seçin.
    </div>
<% End If %>

</div>

<script>
$(function () {
    $('#anagrup').on('change', function () {
        var groupId = $(this).val();
        window.location.href = groupId ? ('join.asp?g=' + encodeURIComponent(groupId)) : 'join.asp';
    });

    $('#tip').on('change', function () {
        var typeId = $(this).val();
        var groupId = '<%=Server.HTMLEncode(grup)%>';
        window.location.href = typeId
            ? ('join.asp?g=' + encodeURIComponent(groupId) + '&tip=' + encodeURIComponent(typeId))
            : ('join.asp?g=' + encodeURIComponent(groupId));
    });

    $('#addSurfaceButton').on('click', function () {
        var name = $.trim($('#isim3').val());

        if (!name) {
            alert('Ürün grubu adını yazın.');
            $('#isim3').focus();
            return;
        }

        $('#surfaceNameHidden').val(name);
        $('#surfaceForm').submit();
    });

    $('#productForm').on('submit', function () {
        var button = $('#productSubmit');
        button.prop('disabled', true).text('Kaydediliyor...');
    });
});
</script>

<br><br>

</body>
