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
Dim grup, tip, surface

grup = Trim(Request.QueryString("g") & "")
tip = Trim(Request.QueryString("tip") & "")
surface = Trim(Request.QueryString("surface") & "")

If grup <> "" And Not IsNumeric(grup) Then grup = ""
If tip <> "" And Not IsNumeric(tip) Then tip = ""
If surface <> "" And Not IsNumeric(surface) Then surface = ""

If Request.QueryString("grup") = "add" Then
  Set Rs = Server.CreateObject("ADODB.RecordSet")
  Sorgula = "SELECT * FROM anagrup"
  Rs.Open Sorgula, baglanti, 1, 3

  Rs.AddNew
  Rs("isim") = Request.Form("isim1")
  Rs("durum") = 1
  Rs.Update

  Rs.Close
  Set Rs = Nothing
  Response.Redirect Request.ServerVariables("HTTP_REFERER")
End If

If Request.QueryString("tip") = "add" Then
  If Not IsNumeric(Request.Form("grup")) Then
    Response.Write "Geçersiz ana kategori."
    Response.End
  End If

  Set Rs = Server.CreateObject("ADODB.RecordSet")
  Sorgula = "SELECT * FROM tip"
  Rs.Open Sorgula, baglanti, 1, 3

  Rs.AddNew
  Rs("isim") = Request.Form("isim2")
  Rs("grup") = CLng(Request.Form("grup"))
  Rs.Update

  Rs.Close
  Set Rs = Nothing
  Response.Redirect Request.ServerVariables("HTTP_REFERER")
End If

If Request.QueryString("surf") = "add" Then
  If Not IsNumeric(Request.Form("grubu")) Then
    Response.Write "Geçersiz kategori."
    Response.End
  End If

  Set Rs = Server.CreateObject("ADODB.RecordSet")
  Sorgula = "SELECT * FROM surface"
  Rs.Open Sorgula, baglanti, 1, 3

  Rs.AddNew
  Rs("isim") = Request.Form("isim3")
  Rs("grup") = CLng(Request.Form("grubu"))
  Rs.Update

  Rs.Close
  Set Rs = Nothing
  Response.Redirect Request.ServerVariables("HTTP_REFERER")
End If
%>

<br><br>

<form name="f" action="join2.asp" method="post">

<table border="1" bordercolor="#DFDFDF" bgcolor="#FFFFFF" cellpadding="5" cellspacing="5" width="900" class="font" align="center">

<script>
$(function () {
  $('#anagrup').on('change', function () {
    var value = $(this).val();
    window.location = value ? ('?g=' + encodeURIComponent(value)) : 'join.asp';
  });

  $('#tip').on('change', function () {
    var value = $(this).val();
    var group = '<%=Server.HTMLEncode(grup)%>';
    window.location = value
      ? ('?g=' + encodeURIComponent(group) + '&tip=' + encodeURIComponent(value))
      : ('?g=' + encodeURIComponent(group));
  });
});
</script>

<tr>
  <td><b>Ürün Kategorisi:</b></td>
  <td>
    <select name="anagrup" id="anagrup">
      <option value="">Lütfen Seçiniz</option>
<%
Set mods = Server.CreateObject("ADODB.RecordSet")
Sorgula = "SELECT * FROM anagrup ORDER BY sira DESC"
mods.Open Sorgula, baglanti, 1, 3

Do While Not mods.EOF
%>
      <option value="<%=mods("id")%>"<%
        If grup <> "" Then
          If CLng(grup) = CLng(mods("id")) Then Response.Write " selected"
        End If
      %>><%=Server.HTMLEncode(mods("isim") & "")%></option>
<%
  mods.MoveNext
Loop
mods.Close
Set mods = Nothing
%>
    </select>
  </td>
  <td colspan="2" width="250">
    <input style="width:150px" name="isim1">
    <input type="submit" value="+ Ana Kategori Ekle" onclick="f.action='join.asp?grup=add'; return true;">
    &nbsp; | &nbsp; <a href="javascript:penc1()"><img src="images/del.png" width="20" alt="Sil"></a>
  </td>
</tr>

<% If grup <> "" Then %>
<tr>
  <td><b>Alt Kategori:</b></td>
  <td>
    <select name="tip" id="tip">
      <option value="">Lütfen Seçiniz</option>
<%
Set mods = Server.CreateObject("ADODB.RecordSet")
Sorgula = "SELECT * FROM tip WHERE grup = " & CLng(grup) & " ORDER BY sira DESC"
mods.Open Sorgula, baglanti, 1, 3

Do While Not mods.EOF
%>
      <option value="<%=mods("id")%>"<%
        If tip <> "" Then
          If CLng(tip) = CLng(mods("id")) Then Response.Write " selected"
        End If
      %>><%=Server.HTMLEncode(mods("isim") & "")%></option>
<%
  mods.MoveNext
Loop
mods.Close
Set mods = Nothing
%>
    </select>
  </td>
  <td colspan="2">
    <input style="width:150px" name="isim2">
    <input type="hidden" name="grup" value="<%=CLng(grup)%>">
    <input type="submit" value="+ Alt Kategori Ekle" onclick="f.action='join.asp?tip=add'; return true;">
    &nbsp; | &nbsp; <a href="javascript:penc2()"><img src="images/del.png" width="20" alt="Sil"></a>
  </td>
</tr>
<% End If %>

<% If grup <> "" And tip <> "" Then %>
<tr>
  <td><b>Ürün Grubu:</b></td>
  <td>
    <select name="surface" id="surface">
      <option value="0">Lütfen Seçiniz</option>
<%
Set mods = Server.CreateObject("ADODB.RecordSet")
Sorgula = "SELECT * FROM surface WHERE grup = " & CLng(grup) & " ORDER BY isim DESC"
mods.Open Sorgula, baglanti, 1, 3

Do While Not mods.EOF
%>
      <option value="<%=mods("id")%>"<%
        If surface <> "" Then
          If CLng(surface) = CLng(mods("id")) Then Response.Write " selected"
        End If
      %>><%=Server.HTMLEncode(mods("isim") & "")%></option>
<%
  mods.MoveNext
Loop
mods.Close
Set mods = Nothing
%>
    </select>
  </td>
  <td colspan="2">
    <input style="width:100px" name="isim3">
    <input type="hidden" name="grubu" value="<%=CLng(grup)%>">
    <input type="submit" value="+ Ekle" onclick="f.action='join.asp?surf=add'; return true;">
    &nbsp; | &nbsp; <a href="javascript:penc3()"><img src="images/del.png" width="20" alt="Sil"></a>
  </td>
</tr>

<tr><td><b>Ürün Kodu:</b></td><td><input style="width:100px" name="kodu" value="0"></td></tr>
<tr><td><b>Ürün Adı:</b></td><td><input style="width:350px" name="isim"></td></tr>

<tr>
  <td><b>Renk Seçenekleri:</b></td>
  <td colspan="3">
<%
Dim renkIndex
For renkIndex = 1 To 10
%>
    <input class="jscolor" name="renk<%=renkIndex%>" value="0" style="width:50px; border-radius:40px; border:1px solid #C0C0C0">
<%
Next
%>
  </td>
</tr>

<tr>
  <td><b>Beden / Numara Seçenekleri:</b></td>
  <td colspan="3">
<%
Dim bedenIndex
For bedenIndex = 1 To 10
%>
    <input name="beden<%=bedenIndex%>" value="0" style="width:40px">
<%
Next
%>
  </td>
</tr>

<tr bgcolor="#F7F4F0">
  <td><b>Özet - Tanımlama:</b></td>
  <td colspan="3"><textarea style="width:700px; height:50px" name="descr"></textarea></td>
</tr>

<%
strFormName = "f"
strTextAreaName = "notlar"
%>

<tr bgcolor="#F7F4F0">
  <td colspan="4">
    <b>Ürün Detayları:</b><br><br>
    <!--#include file="RTE_editor_inc.asp"--><br>
    <textarea style="width:870px; height:390px" name="notlar" id="notlar"></textarea>
  </td>
</tr>

<tr><td><b>Kargoya Veriliş Süresi:</b></td><td><input style="width:100px" name="delivery" value="3"></td></tr>

<tr>
  <td><b>Anahtar Kelimeler:</b></td>
  <td colspan="3"><textarea style="width:700px; height:50px" name="keyw"></textarea></td>
</tr>

<input type="hidden" name="yayin" value="0">

<tr>
  <td colspan="4" align="center"><input type="submit" value="ÜRÜN EKLE"></td>
</tr>

<% End If %>

</table>
</form>

<br><br>

</body>
