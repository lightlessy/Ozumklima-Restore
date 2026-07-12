<%@ Language="VBScript" CodePage="65001" %>
<!--#include file='charset.asp'-->

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html>
<head>

<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">

<title>Özüm Klima İletişim | Antalya Klima Servisi ve Keşif Talebi</title>

<meta name="description" content="Özüm Klima ile iletişime geçin. Antalya'da klima, VRF, havalandırma, keşif, teklif, montaj, bakım ve servis desteği alın.">
<link rel="canonical" href="https://ozumklima.com/iletisim.asp">

<meta name="author" content="Mega Tasarım">

<link rel="icon" type="image/png" href="fav.png">

<link href="css/font.css" rel="stylesheet" type="text/css">

<script src="js/jquery-2.1.1.js"></script>
<script src="js/jquery-ui.min.js"></script>

<link href="css/fontawesome-all.css" rel="stylesheet">

<%
hform = Request.QueryString("hform")
%>

</head>

<body>

<!--#include file='ust.asp'-->

<div>

  <div class="slogan" style="height:20px"></div>

  <div class="iletTas">

    <div style="float:left; max-width:300px" class="adresTel">

      <span style="font-size:1.3em; color:#FEAD41;">Alo Konfor Hattı<br></span>

      <span class="alttel">
        <i class="fas fa-phone"></i>
        <%=ayars("tel")%>
        <br><br>
      </span>

      <strong>Adres:</strong><br>
      <%=ayars("adres")%>

      <br><br>

      <strong>E-Posta:</strong><br>
      <a href="mailto:<%=ayars("eposta")%>" target="_blank" style="color:#000000">
        <%=ayars("eposta")%>
      </a>

      <br><br>

      <strong>Faks:</strong><br>
      <%=ayars("faks")%>

      <br><br>

      <strong>Bizi Takip Edin:</strong><br>

      <a href="<%=ayars("face")%>" target="_blank" rel="noopener">
        <i class="fab fa-facebook-f"></i>
      </a>
      &nbsp;&nbsp;

      <a href="<%=ayars("insta")%>" target="_blank" rel="noopener">
        <i class="fab fa-instagram"></i>
      </a>

    </div>

    <div style="float:right;" class="formTas">

      <div style="text-align:center;">

        <form
          method="post"
          name="cic"
          action="iletisim-send.asp"
          class="form"
          onsubmit="safeTrack('contact_form_submit_attempt', {source:'iletisim_page'});"
        >

          <input
            name="isim"
            type="text"
            placeholder="İsim"
            size="26"
            maxlength="60"
            required
          >

          <input
            name="mail"
            type="email"
            placeholder="E-Posta Adresiniz"
            size="26"
            maxlength="60"
            required
          >

          <input
            name="tel"
            type="text"
            placeholder="Telefon Numaranız"
            size="10"
            maxlength="11"
            onkeypress="return numbersonly(this, event)"
          >

          <textarea
            name="mes"
            cols="25"
            rows="6"
            placeholder="Mesajınız"
            required
          ></textarea>

          <input
            type="submit"
            name="submit"
            value="Gönder"
            class="btn"
            style="color:#FFFFFF"
          >

        </form>

      </div>

    </div>

    <div style="clear:both"></div>

  </div>

</div>

<%=ayars("map")%>

<!--#include file='alt.asp'-->

</body>
</html>