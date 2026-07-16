<%
Dim ozmProductCardPage
ozmProductCardPage = LCase(Request.ServerVariables("SCRIPT_NAME"))
If InStr(ozmProductCardPage, "/klima.asp") > 0 Or _
   InStr(ozmProductCardPage, "/vrf_sistemler.asp") > 0 Or _
   InStr(ozmProductCardPage, "/isi_pompasi.asp") > 0 Or _
   InStr(ozmProductCardPage, "/havalandirma.asp") > 0 Then
%>
<link rel="stylesheet" href="css/product-cards.css">
<%
End If
%>

<!--#include file='db.asp'-->

<div class="menuLogoTas">

  <div class="ozm-topbar">
    <div class="ozm-topbar-inner">
      <div class="ozm-topbar-text">
        Klima, VRF, Is&#305; Pompas&#305; ve Havaland&#305;rma &#199;&#246;z&#252;mleri
      </div>

      <div class="ozm-topbar-links">
        <a href="tel:02422297139" onclick="safeTrack('phone_click', {location:'top_bar'});">
          0242 229 71 39
        </a>

        <a href="iletisim.asp" onclick="safeTrack('lead_click', {location:'top_bar'});">
          Ke&#351;if Talebi
        </a>
      </div>
    </div>
  </div>

  <div class="ozm-header-main">

    <div id="nav-icon2" class="ackapa" aria-label="Men&#252;y&#252; a&#231; veya kapat">
      <span></span>
      <span></span>
      <span></span>
      <span></span>
      <span></span>
      <span></span>
    </div>

    <div class="ozm-logo-wrap">
      <a href="https://ozumklima.com" aria-label="&#214;z&#252;m Klima Ana Sayfa">
<img src="/images/logo.png" class="logo ozm-logo" alt="&#214;z&#252;m Klima">
      </a>
    </div>

    <div class="ozm-search-wrap">
      <form method="POST" action="bul.asp" name="aform" class="ozm-search-form">
        <input type="text" name="bul" required class="ozm-search-input" placeholder="Klima, VRF, &#305;s&#305; pompas&#305; ara">
        <button type="submit" class="ozm-search-button" aria-label="Ara">&#128269;</button>
      </form>
    </div>

    <div class="ozm-header-actions">
      <a class="ozm-call-btn" href="tel:02422297139" onclick="safeTrack('phone_click', {location:'header'});">
        <i class="fas fa-phone-square"></i>
        <span>Ara</span>
      </a>

      <a class="ozm-cta-btn" href="iletisim.asp" onclick="safeTrack('lead_click', {location:'header'});">
        Ke&#351;if Al
      </a>
    </div>

  </div>

  <div style="clear:both;"></div>

  <div class="serit">
    <div class="menu">

      <a href="h.asp?id=12">
        <div>H&#304;ZMETLER</div>
      </a>
      <a href="uygulama-ornekleri.asp">
        <div>UYGULAMA ORNEKLERI</div>
      </a>
      <a href="klima.asp">
        <div>KL&#304;MA</div>
      </a>
      <a href="havalandirma.asp">
        <div>HAVALANDIRMA S&#304;STEMLER&#304;</div>
      </a>
      <a href="vrf_sistemler.asp">
        <div>VRF S&#304;STEM</div>
      </a>
      <a href="mitsubishi-electric-isi-pompasi-(heat-pump)-165">
        <div>ISI POMPASI</div>
      </a>
      <a href="galeri.asp">
        <div>GALER&#304;</div>
      </a>
      <a href="iletisim.asp">
        <div>&#304;LET&#304;&#350;&#304;M</div>
      </a>
    </div>
  </div>

</div>

<style>
  .ozm-topbar,
  .ozm-topbar *,
  .ozm-header-main,
  .ozm-header-main * {
    box-sizing: border-box;
  }

  .ozm-topbar {
    width: 100%;
    background: #f5f5f5;
    border-bottom: 1px solid #e8e8e8;
    color: #333;
    font-size: 13px;
  }

  .ozm-topbar-inner {
    max-width: 1220px;
    min-height: 36px;
    margin: 0 auto;
    padding: 0 18px;
    display: flex;
    align-items: center;
    justify-content: space-between;
    gap: 16px;
  }

  .ozm-topbar-text {
    font-weight: 500;
  }

  .ozm-topbar-links {
    display: flex;
    align-items: center;
    gap: 16px;
    white-space: nowrap;
  }

  .ozm-topbar-links a {
    color: #222;
    text-decoration: none;
    font-weight: 700;
  }

  .ozm-header-main {
    max-width: 1220px;
    min-height: 92px;
    margin: 0 auto;
    padding: 16px 18px;
    display: flex;
    align-items: center;
    gap: 22px;
  }

  .ozm-logo-wrap {
    width: 245px;
    flex: 0 0 245px;
  }

  .ozm-logo {
    display: block;
    width: 245px;
    max-width: 100%;
    height: auto;
  }

  .ozm-search-wrap {
    flex: 1;
    max-width: 390px;
  }

  .ozm-search-form {
    width: 100%;
    display: flex;
    align-items: center;
    border-bottom: 1px solid #e5e5e5;
  }

  .ozm-search-input {
    width: 100%;
    height: 40px;
    border: 0;
    outline: 0;
    background: transparent;
    font-size: 14px;
  }

  .ozm-search-button {
    border: 0;
    background: transparent;
    cursor: pointer;
    font-size: 18px;
    line-height: 1;
  }

  .ozm-header-actions {
    margin-left: auto;
    display: flex;
    align-items: center;
    gap: 10px;
    white-space: nowrap;
  }

  .ozm-call-btn,
  .ozm-cta-btn {
    min-height: 42px;
    padding: 0 15px;
    display: inline-flex;
    align-items: center;
    justify-content: center;
    gap: 7px;
    border-radius: 4px;
    text-decoration: none;
    font-size: 14px;
    font-weight: 800;
  }

  .ozm-call-btn {
    color: #222;
    background: #fff;
    border: 1px solid #ddd;
  }

  .ozm-cta-btn {
    color: #fff;
    background: #222;
    border: 1px solid #222;
  }

  #nav-icon2 {
    width: 30px;
    height: 35px;
    margin-right: 5px;
    margin-top: 3px;
    z-index: 999;
    transform: rotate(0deg);
    transition: .5s ease-in-out;
    cursor: pointer;
    position: relative;
    flex: 0 0 30px;
    min-width: 30px;
  }

  #nav-icon2 span {
    display: block;
    position: absolute;
    height: 4px;
    width: 50%;
    background: #3B3B3B;
    opacity: 1;
    top: 0px;
    transform: rotate(0deg);
    transition: .25s ease-in-out;
  }

  #nav-icon2 span:nth-child(even) {
    left: 50%;
    border-radius: 0 9px 9px 0;
  }

  #nav-icon2 span:nth-child(odd) {
    left: 0px;
    border-radius: 9px 0 0 9px;
  }

  #nav-icon2 span:nth-child(1),
  #nav-icon2 span:nth-child(2) {
    top: 0px;
  }

  #nav-icon2 span:nth-child(3),
  #nav-icon2 span:nth-child(4) {
    top: 9px;
  }

  #nav-icon2 span:nth-child(5),
  #nav-icon2 span:nth-child(6) {
    top: 18px;
  }

  #nav-icon2.open span:nth-child(1),
  #nav-icon2.open span:nth-child(6) {
    transform: rotate(45deg);
  }

  #nav-icon2.open span:nth-child(2),
  #nav-icon2.open span:nth-child(5) {
    transform: rotate(-45deg);
  }

  #nav-icon2.open span:nth-child(1) {
    left: 5px;
    top: 7px;
  }

  #nav-icon2.open span:nth-child(2) {
    left: calc(50% - 7px);
    top: 7px;
  }

  #nav-icon2.open span:nth-child(3) {
    left: -50%;
    opacity: 0;
  }

  #nav-icon2.open span:nth-child(4) {
    left: 100%;
    opacity: 0;
  }

  #nav-icon2.open span:nth-child(5) {
    left: 4px;
    top: 18px;
  }

  #nav-icon2.open span:nth-child(6) {
    left: calc(50% - 0px);
    top: 18px;
  }

  @media screen and (min-width: 1025px) {
    #nav-icon2 {
      display: none;
    }
  }

  @media screen and (max-width: 1024px) {
    .ozm-topbar {
      display: none;
    }

    .ozm-header-main {
      min-height: 70px;
      padding: 10px 14px;
      gap: 12px;
    }

    .ozm-logo-wrap {
      width: 160px;
      flex: 0 0 160px;
    }

    .ozm-logo {
      width: 160px;
    }

    .ozm-search-wrap {
      display: none;
    }

    .ozm-header-actions {
      margin-left: auto;
    }

    .ozm-call-btn {
      width: 42px;
      height: 42px;
      padding: 0;
      border-radius: 50%;
    }

    .ozm-call-btn span {
      display: none;
    }

    .ozm-cta-btn {
      min-height: 38px;
      padding: 0 12px;
      font-size: 13px;
    }
  }

  @media screen and (max-width: 420px) {
    .ozm-logo-wrap {
      width: 140px;
      flex: 0 0 140px;
    }

    .ozm-logo {
      width: 140px;
    }

    .ozm-cta-btn {
      display: none;
    }
  }
  @media screen and (max-width: 520px) {
  .ozm-cta-btn {
    display: none;
  }

  .ozm-header-main {
    gap: 10px;
  }

  .ozm-logo-wrap {
    width: 140px;
    flex: 0 0 140px;
  }

  .ozm-logo {
    width: 140px;
  }
}
</style>

<script>
  function safeTrack(eventName, params) {
    if (typeof analyticsTrack === "function") {
      analyticsTrack(eventName, params || {});
    }
  }

  $(document).ready(function () {
    var acikmi = false;

    $('#nav-icon2').click(function () {
      $(this).toggleClass('open');
      acikmi = !acikmi;

      if (acikmi) {
        $(".menu").slideDown();
      } else {
        $(".menu").slideUp();
      }
    });
  });
</script>

<script>
  $(document).ready(function () {
    var ekranGen = $(window).width();

    if (ekranGen > 1024) {
      var didScroll = false;

      window.onscroll = function () {
        didScroll = true;
      };

      setInterval(function () {
        if (didScroll) {
          didScroll = false;

          var scroll = $(window).scrollTop();

          if (scroll > 80) {
            $(".logo").attr("src", "images/logo2.png");
            $(".logo").stop().animate({ width: "200px" }, 200);
            $(".menuLogoTas").stop().animate({ paddingTop: "0px" }, 200);
            $(".genislik").stop().animate({ width: "250px" }, 200);
          } else {
            $(".logo").attr("src", "images/logo.png");
            $(".logo").stop().animate({ width: "245px" }, 200);
            $(".menuLogoTas").stop().animate({ paddingTop: "0px" }, 200);
            $(".genislik").stop().animate({ width: "300px" }, 200);
          }
        }
      }, 100);
    } else {
      $(".logo").attr("src", "images/logo2.png");
    }

    if (ekranGen < 1024) {
      $(".MenuVeicerik").before().click(function () {
        $(".solMenu").slideToggle();
      });
    }
  });
</script>