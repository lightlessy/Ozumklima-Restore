<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html>
<head>

<META HTTP-EQUIV="Content-Type" content="text/html; charset=iso-8859-9">
<meta name="viewport" content="width=device-width, initial-scale=1">

<title>Tesekkurler | Antalya Klima - Ozum Klima Mitsubishi Klima Antalya</title>

<meta name="author" content="Mega Tasarim"/>

<link rel="icon" type="image/png" href="fav.png" />
<link href="css/font.css" rel="stylesheet" type="text/css"/>
<link href="css/fontawesome-all.css" rel="stylesheet">

</head>

<body>

<!--#include file='ust.asp'-->

<DIV>
  <Div class="slogan" style="height:20px"></Div>

  <Div class="iletTas" style="padding:20px 10px 40px 10px;">
    <div style="max-width:900px; margin:0px auto; border:1px solid #E4E4E4; background-color:#FFFFFF; padding:30px; font-family:'Source Sans Pro', sans-serif; color:#2E2E2E;">
      <div style="font-size:1.8em; font-weight:700; color:#3E3E3E; margin-bottom:15px;">Tesekkurler</div>

      <div style="font-size:1.2em; line-height:1.6em; color:#3E3E3E;">
        <b>Talebiniz alinmistir. Ozum Klima ekibi en kisa surede sizinle iletisime gececektir.</b>
      </div>

      <div style="margin-top:24px;">
        <a href="tel:<%=Replace(ayars("tel")," ","")%>" style="display:inline-block; margin:0 10px 10px 0; padding:10px 16px; border:1px solid #D7D7D7; background-color:#F4F4F4; color:#3E3E3E; font-weight:700;" onclick="analyticsTrack('phone_click', {location: 'tesekkurler'});">Hemen ara</a>
        <a href="https://api.whatsapp.com/send?phone=9<%=Replace(ayars("gsm")," ","")%>&text=https://www.ozumklima.com" target="_blank" style="display:inline-block; margin:0 10px 10px 0; padding:10px 16px; border:1px solid #D7D7D7; background-color:#F4F4F4; color:#3E3E3E; font-weight:700;" onclick="analyticsTrack('whatsapp_click', {location: 'tesekkurler'});">WhatsApp'tan yaz</a>
      </div>

      <div id="sayacMetni" style="margin-top:12px; color:#5F5F5F; font-size:1em;">
        <span id="kalanSure">10</span> saniye sonra ana sayfaya yonlendirileceksiniz.
      </div>

      <div style="margin-top:8px;">
        <a href="https://ozumklima.com" style="color:#3E3E3E; font-weight:700;">Ana sayfaya simdi don</a>
      </div>
    </div>
  </Div>
</DIV>

<!--  #include file='alt.asp' -->

<script>
  analyticsTrack('contact_form_submit', {status: 'success', source: 'iletisim_formu'});

  var kalan = 10;
  var sayac = document.getElementById("kalanSure");
  var sayacMetni = document.getElementById("sayacMetni");

  function metinGuncelle() {
    if (kalan === 1) {
      sayacMetni.innerHTML = "<span id='kalanSure'>" + kalan + "</span> saniye sonra ana sayfaya yonlendirileceksiniz.";
    } else {
      sayacMetni.innerHTML = "<span id='kalanSure'>" + kalan + "</span> saniye sonra ana sayfaya yonlendirileceksiniz.";
    }
    sayac = document.getElementById("kalanSure");
  }

  metinGuncelle();

  var timer = setInterval(function () {
    kalan = kalan - 1;

    if (kalan >= 0) {
      metinGuncelle();
    }

    if (kalan <= 0) {
      clearInterval(timer);
      window.location.href = "https://ozumklima.com";
    }
  }, 1000);
</script>

</body>
</html>
