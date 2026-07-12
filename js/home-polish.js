<script>
(function () {
  function isHomePage() {
    var path = (window.location.pathname || "").toLowerCase();
    return path === "/" || path === "" || path.indexOf("index.asp") !== -1;
  }

  function addHomePolish() {
    if (!isHomePage()) {
      return;
    }

    document.body.classList.add("ozm-home-polish-active");

    var slogan = document.querySelector(".slogan");
    if (slogan && !document.querySelector(".homeCta")) {
      slogan.insertAdjacentHTML(
        "beforeend",
        '<div class="homeCta">' +
          '<a href="iletisim.asp" onclick="safeTrack(\'lead_click\', {location:\'home_slogan\', action:\'kesif_talep\'});">Ke&#351;if Talep Et</a>' +
          '<a href="https://api.whatsapp.com/send?phone=902422297139&text=https://www.ozumklima.com" target="_blank" onclick="safeTrack(\'whatsapp_click\', {location:\'home_slogan\'});">WhatsApp\'tan Yaz</a>' +
        '</div>'
      );
    }

    var referenceTitle = document.querySelector(".anlasmali .slo1 span");
    if (referenceTitle && referenceTitle.innerHTML.indexOf("g&uuml;ven") === -1) {
      referenceTitle.innerHTML = "Referanslar";
    }
  }

  if (document.readyState === "loading") {
    document.addEventListener("DOMContentLoaded", addHomePolish);
  } else {
    addHomePolish();
  }
})();
</script>
