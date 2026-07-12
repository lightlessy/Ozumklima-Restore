<%@ Language="VBScript" CodePage="65001" %>
<%
Response.Buffer = True
Response.CodePage = 65001
Response.Charset = "utf-8"
%>

  <!--#include file="menu.asp"-->

  <body onload="initialiseWebWizRTE();">

  <div class="baslik">ÃœrÃ¼n Ekle</div>

  <script src="../js/jscolor.js"></script>

  <%
  Dim grup, tip, surface
  Dim Rs, mods, Sorgula

  grup = Trim(Request.QueryString("g") & "")
  tip = Trim(Request.QueryString("tip") & "")
  surface = Trim(Request.QueryString("surface") & "")

  If Not IsNumeric(grup) Then grup = ""
  If Not IsNumeric(tip) Then tip = ""
  If Not IsNumeric(surface) Then surface = ""

  If Request.QueryString("grup") = "add" Then
  Set Rs = Server.CreateObject("ADODB.Recordset")
  Sorgula = "SELECT * FROM anagrup"
  Rs.Open Sorgula, baglanti, 1, 3

  Rs.AddNew
  Rs("isim") = Trim(Request.Form("isim1") & "")
  Rs("durum") = 1
  Rs.Update

  Rs.Close
  Set Rs = Nothing

  Response.Redirect Request.ServerVariables("HTTP_REFERER")
  Response.End
  End If

  If Request.QueryString("tip") = "add" Then
  If IsNumeric(Request.Form("grup")) Then

      Set Rs = Server.CreateObject("ADODB.Recordset")
      Sorgula = "SELECT * FROM tip"
      Rs.Open Sorgula, baglanti, 1, 3

      Rs.AddNew
      Rs("isim") = Trim(Request.Form("isim2") & "")
      Rs("grup") = CLng(Request.Form("grup"))
      Rs.Update

      Rs.Close
      Set Rs = Nothing

  End If

  Response.Redirect Request.ServerVariables("HTTP_REFERER")
  Response.End
  End If

  If Request.QueryString("surf") = "add" Then
  If IsNumeric(Request.Form("grubu")) Then

      Set Rs = Server.CreateObject("ADODB.Recordset")
      Sorgula = "SELECT * FROM surface"
      Rs.Open Sorgula, baglanti, 1, 3

      Rs.AddNew
      Rs("isim") = Trim(Request.Form("isim3") & "")
      Rs("grup") = CLng(Request.Form("grubu"))
      Rs.Update

      Rs.Close
      Set Rs = Nothing

  End If

  Response.Redirect Request.ServerVariables("HTTP_REFERER")
  Response.End
  End If
  %>

  <br><br>

  <form name="f" action="join2.asp" method="post">

  <table
      border="1"
      bordercolor="#DFDFDF"
      bgcolor="#FFFFFF"
      cellpadding="5"
      cellspacing="5"
      width="900"
      class="font"
      align="center"
  >

  <script>
  $(function () {
      $('#anagrup').on('change', function () {
          var selectedGroup = $(this).val();

          if (selectedGroup !== '') {
              window.location.href = 'join.asp?g=' + encodeURIComponent(selectedGroup);
          } else {
              window.location.href = 'join.asp';
          }
      });

      $('#tip').on('change', function () {
          var selectedType = $(this).val();
          var selectedGroup = '<%=Server.HTMLEncode(grup)%>';

          if (selectedType !== '' && selectedGroup !== '') {
              window.location.href =
                  'join.asp?g=' +
                  encodeURIComponent(selectedGroup) +
                  '&tip=' +
                  encodeURIComponent(selectedType);
          }
      });
  });
  </script>

  <tr>
      <td><strong>ÃœrÃ¼n Kategori:</strong></td>
  <td>
      <select name="anagrup" id="anagrup">
          <option value="">LÃ¼tfen SeÃ§iniz</option>

          <%
          Set mods = Server.CreateObject("ADODB.Recordset")
          Sorgula = "SELECT * FROM anagrup ORDER BY sira DESC"
          mods.Open Sorgula, baglanti, 1, 3

          Do While Not mods.EOF
          %>

              <option value="<%=CLng(mods("id"))%>"<%
                  If grup <> "" Then
                      If CLng(grup) = CLng(mods("id")) Then
                          Response.Write " selected"
                      End If
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
      <input type="text" style="width:150px" name="isim1">

      <input
          type="submit"
          value="+ Ana Kategori Ekle"
          onclick="f.action='join.asp?grup=add'; return true;"
      >

      &nbsp; | &nbsp;

      <a href="javascript:penc1()">
          <img src="images/del.png" width="20" alt="Sil">
      </a>
  </td>
  </tr>

  <%
  If grup <> "" Then
  %>

  <tr>
      <td><strong>Alt Kategori:</strong></td>
  <td>
      <select name="tip" id="tip">
          <option value="">LÃ¼tfen SeÃ§iniz</option>

          <%
          Set mods = Server.CreateObject("ADODB.Recordset")
          Sorgula = "SELECT * FROM tip WHERE grup = " & CLng(grup) & " ORDER BY sira DESC"
          mods.Open Sorgula, baglanti, 1, 3

          Do While Not mods.EOF
          %>

              <option value="<%=CLng(mods("id"))%>"<%
                  If tip <> "" Then
                      If CLng(tip) = CLng(mods("id")) Then
                          Response.Write " selected"
                      End If
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
      <input type="text" style="width:150px" name="isim2">

      <input type="hidden" name="grup" value="<%=Server.HTMLEncode(grup)%>">

      <input
          type="submit"
          value="+ Alt Kategori Ekle"
          onclick="f.action='join.asp?tip=add'; return true;"
      >

      &nbsp; | &nbsp;

      <a href="javascript:penc2()">
          <img src="images/del.png" width="20" alt="Sil">
      </a>
  </td>
  </tr>

  <%
  End If
  %>

  <%
  If tip <> "" And grup <> "" Then
  %>

  <tr>
      <td><strong>ÃœrÃ¼n Grubu:</strong></td>
  <td>
      <select name="surface" id="surface">
          <option value="0">LÃ¼tfen SeÃ§iniz</option>

          <%
          Set mods = Server.CreateObject("ADODB.Recordset")
          Sorgula = "SELECT * FROM surface WHERE grup = " & CLng(grup) & " ORDER BY isim DESC"
          mods.Open Sorgula, baglanti, 1, 3

          Do While Not mods.EOF
          %>

              <option value="<%=CLng(mods("id"))%>"<%
                  If surface <> "" Then
                      If CLng(surface) = CLng(mods("id")) Then
                          Response.Write " selected"
                      End If
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
      <input type="text" style="width:100px" name="isim3">

      <input type="hidden" name="grubu" value="<%=Server.HTMLEncode(grup)%>">

      <input
          type="submit"
          value="+ Ekle"
          onclick="f.action='join.asp?surf=add'; return true;"
      >

      &nbsp; | &nbsp;

      <a href="javascript:penc3()">
          <img src="images/del.png" width="20" alt="Sil">
      </a>
  </td>
  </tr>

  <tr>
      <td><strong>ÃœrÃ¼n Kodu:</strong></td>
      <td>
          <input type="text" style="width:100px" name="kodu" value="0">
      </td>
  </tr>

  <tr>
      <td><strong>ÃœrÃ¼n AdÄ±:</strong></td>
      <td>
          <input type="text" style="width:350px" name="isim">
      </td>
  </tr>

  <tr>
      <td><strong>Renk SeÃ§enekleri:</strong></td>
  <td colspan="3">
      <input class="jscolor" name="renk1" value="0" style="width:50px; border-radius:40px; border:1px solid #C0C0C0">
      <input class="jscolor" name="renk2" value="0" style="width:50px; border-radius:40px; border:1px solid #C0C0C0">
      <input class="jscolor" name="renk3" value="0" style="width:50px; border-radius:40px; border:1px solid #C0C0C0">
      <input class="jscolor" name="renk4" value="0" style="width:50px; border-radius:40px; border:1px solid #C0C0C0">
      <input class="jscolor" name="renk5" value="0" style="width:50px; border-radius:40px; border:1px solid #C0C0C0">
      <input class="jscolor" name="renk6" value="0" style="width:50px; border-radius:40px; border:1px solid #C0C0C0">
      <input class="jscolor" name="renk7" value="0" style="width:50px; border-radius:40px; border:1px solid #C0C0C0">
      <input class="jscolor" name="renk8" value="0" style="width:50px; border-radius:40px; border:1px solid #C0C0C0">
      <input class="jscolor" name="renk9" value="0" style="width:50px; border-radius:40px; border:1px solid #C0C0C0">
      <input class="jscolor" name="renk10" value="0" style="width:50px; border-radius:40px; border:1px solid #C0C0C0">
  </td>
  </tr>

  <tr>
      <td><strong>Beden / Numara SeÃ§enekleri:</strong></td>
  <td colspan="3">
      <input name="beden1" value="0" style="width:40px">
      <input name="beden2" value="0" style="width:40px">
      <input name="beden3" value="0" style="width:40px">
      <input name="beden4" value="0" style="width:40px">
      <input name="beden5" value="0" style="width:40px">
      <input name="beden6" value="0" style="width:40px">
      <input name="beden7" value="0" style="width:40px">
      <input name="beden8" value="0" style="width:40px">
      <input name="beden9" value="0" style="width:40px">
      <input name="beden10" value="0" style="width:40px">
  </td>
  </tr>

  <tr bgcolor="#F7F4F0">
      <td><strong>Ã–zet - TanÄ±mlama:</strong></td>
  <td colspan="3">
      <textarea
          style="width:700px; height:50px"
          name="descr"
      ></textarea>
  </td>
  </tr>

  <%
  strFormName = "f"
  strTextAreaName = "notlar"
  %>

  <tr bgcolor="#F7F4F0">
      <td colspan="4">
          <strong>ÃœrÃ¼n DetaylarÄ±:</strong>
  <br><br>

      <!--#include file="RTE_editor_inc.asp"-->

      <br>

      <textarea
          style="width:870px; height:390px"
          name="notlar"
          id="notlar"
      ></textarea>
  </td>
  </tr>

  <tr>
      <td><strong>Kargoya VeriliÅŸ SÃ¼resi:</strong></td>
  <td>
      <input
          type="text"
          style="width:100px"
          name="delivery"
          value="3"
      >
  </td>
  </tr>

  <tr>
      <td><strong>Anahtar Kelimeler:</strong></td>
  <td colspan="3">
      <textarea
          style="width:700px; height:50px"
          name="keyw"
      ></textarea>
  </td>
  </tr>

  <input type="hidden" name="yayin" value="0">

  <tr>
      <td colspan="4" align="center">
          <input type="submit" value="ÃœRÃœN EKLE">
      </td>
  </tr>

  <%
  End If
  %>

  </table>

  </form>

  <br><br>

  </body>

